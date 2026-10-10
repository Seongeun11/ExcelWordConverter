#model.py
import os
import json
import re
from pathlib import Path
import pandas as pd
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

class ExcelModel:
    def __init__(self):
        pass
        #self.min_length = 1

    def clean_text(self, text: str) -> str:
        text = str(text).strip()
        text = re.sub(r"\s+", " ", text)
        return text

    def process_excel_to_json(self, excel_path: Path, target_cols: list, json_output_dir: Path, max_bytes=20 * 1024,min_length=30):
        df = pd.read_excel(excel_path)
        original_columns = list(df.columns)
        df.columns = list(range(len(df.columns)))
        df = df.fillna('')

        name_col = 1   
        church_col = 4 

        if church_col in df.columns:
            df[church_col] = (
                df[church_col]
                .astype(str)
                .str.strip()
                .str.replace('교회', '', regex=False)
                .str.strip()
            )
            df[church_col] = df[church_col].apply(
                lambda x: x + '교회' if x != '' else ''
            )

        json_output_dir.mkdir(parents=True, exist_ok=True)
        generated_json_files = []

        for ngtv_num_index in target_cols:
            col_idx = ngtv_num_index - 1
            if col_idx < 0 or col_idx >= len(df.columns):
                continue

            result = []
            i = 0
            for _, row in df.iterrows():
                answer = self.clean_text(row[col_idx])
                if len(answer) < min_length:
                    continue
                
                i += 1
                entry = {
                    "번호": i,
                    "이름": str(row[name_col]).strip() if name_col in df.columns else "",
                    "소속교회": str(row[church_col]).strip() if church_col in df.columns else "",
                    "응답": answer
                }
                result.append(entry)

            if not result:
                continue

            original_col_name = str(original_columns[col_idx])
            safe_filename = (
                original_col_name
                .replace(' ', '_')
                .replace('?', '')
                .replace('/', '')
                .replace('\n', '_')
            )

            current_chunk = []
            part_num = 1

            for item in result:
                test_chunk = current_chunk + [item]
                json_bytes = len(json.dumps(test_chunk, ensure_ascii=False, indent=2).encode('utf-8'))

                if json_bytes > max_bytes and current_chunk:
                    part_path = json_output_dir / f"{safe_filename}_part_{part_num}.json"
                    with open(part_path, 'w', encoding='utf-8') as f:
                        json.dump(current_chunk, f, ensure_ascii=False, indent=2)
                    generated_json_files.append(part_path)

                    current_chunk = [item]
                    part_num += 1
                else:
                    current_chunk = test_chunk

            if current_chunk:
                if part_num == 1:
                    part_path = json_output_dir / f"{safe_filename}.json"
                else:
                    part_path = json_output_dir / f"{safe_filename}_part_{part_num}.json"

                with open(part_path, 'w', encoding='utf-8') as f:
                    json.dump(current_chunk, f, ensure_ascii=False, indent=2)
                generated_json_files.append(part_path)

        return generated_json_files

    def set_paragraph_border(self, paragraph, color="203A7B", size="20"):
        p = paragraph._p
        pPr = p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        for pos in ['top', 'bottom']:
            el = OxmlElement(f'w:{pos}')
            el.set(qn('w:val'), 'single')
            el.set(qn('w:sz'), size)
            el.set(qn('w:space'), '1')
            el.set(qn('w:color'), color)
            pBdr.append(el)
        pPr.append(pBdr)

    def add_title(self, doc, text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(16)
        run.font.color.rgb = RGBColor(27, 23, 96)
        self.set_paragraph_border(p)
        return p

    def get_grouped_json_files(self, sub_dir: Path):
        """
        [핵심 로직] 
        동일한 원본 파일명을 가진 파일들과 분할 파일들(_part_N)을 그룹화하고, 
        파트 번호 기준 오름차순(숫자 크기순)으로 정렬합니다.
        """
        json_files = list(sub_dir.glob("*.json"))
        groups = {}

        for jf in json_files:
            stem = jf.stem
            if "_part_" in stem:
                base_name, part_str = stem.split("_part_", 1)
                try:
                    part_num = int(part_str)
                except ValueError:
                    part_num = 0
            else:
                base_name = stem
                part_num = 0  # 분할되지 않은 원본 파일은 0순위(가장 먼저 배치)

            groups.setdefault(base_name, []).append((part_num, jf))

        # 파트 번호 숫자를 기준으로 정확히 정렬 (예: part_2 다음 part_10 보장)
        sorted_groups = {}
        for base_name, file_tuples in groups.items():
            file_tuples.sort(key=lambda x: x[0])
            sorted_groups[base_name] = [jf for _, jf in file_tuples]

        return sorted_groups

    def convert_json_to_word_from_directory(self, json_output_root: Path, output_word_dir: Path, logger_callback=None):
        """
        디렉터리 내의 모든 JSON 그룹을 탐색하여 하나의 Word 파일로 병합 변환합니다.
        """
        output_word_dir.mkdir(parents=True, exist_ok=True)
        sub_dirs = [d for d in json_output_root.iterdir() if d.is_dir()]
        
        if not sub_dirs:
            if logger_callback:
                logger_callback("⚠️ 변환할 JSON 하위 폴더가 없습니다.")
            return []

        saved_files = []
        for sub_dir in sub_dirs:
            original_stem = sub_dir.name
            if logger_callback:
                logger_callback(f"  └─ 처리 중: {original_stem}")

            question_groups = self.get_grouped_json_files(sub_dir)
            if not question_groups:
                continue

            for q_name, q_files in question_groups.items():
                output_docx = self.generate_word_document(q_files, output_word_dir, q_name)
                saved_files.append(output_docx)
                if logger_callback:
                    logger_callback(f"     └─ Word 저장 완료: {output_docx.name}")

        return saved_files

    def generate_word_document(self, json_files: list, output_word_dir: Path, original_stem: str):
        doc = Document()
        title_text = original_stem.replace('_', ' ')
        self.add_title(doc, title_text)
        doc.add_paragraph("")

        for json_file in json_files:
            if not Path(json_file).exists():
                continue
            with open(json_file, 'r', encoding='utf-8') as f:
                try:
                    data = json.load(f)
                except json.JSONDecodeError:
                    continue

            for item in data:
                number = item.get("번호", "")
                name = item.get("이름", "")
                church = item.get("소속교회", "")
                content = item.get("응답", "")

                p1 = doc.add_paragraph()
                run_num = p1.add_run(f"{number}. ")
                run_num.font.color.rgb = RGBColor(255, 0, 0)
                run_num.font.size = Pt(12)

                run_title = p1.add_run(f"< {name} / {church} >")
                run_title.font.color.rgb = RGBColor(0, 0, 255)
                run_title.font.size = Pt(12)
                run_title.bold = True

                p2 = doc.add_paragraph()
                run_content = p2.add_run(content)
                run_content.font.size = Pt(11)

                doc.add_paragraph("")

        output_file_path = output_word_dir / f"{original_stem}_word.docx"
        doc.save(str(output_file_path))
        return output_file_path