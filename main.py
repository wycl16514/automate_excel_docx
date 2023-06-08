import glob
import os.path
from pathlib import Path

from excelReader import ExcelReader
from wordWriter import WordWriter

# word_writer = WordWriter('/Users/my/Downloads/special_building_template.docx', 'modified.docx')
# word_writer.set_name("陈毅")
# word_writer.save()
# excel_reader = ExcelReader('/Users/my/Downloads/建筑报名表生成专用.xlsx')
# print(excel_reader.get_content(1, "身份证号"))

import tempfile
import zipfile
from excelReader import ExcelReader
from wordWriter import WordWriter


def extract_xlsx_docx(xlsx_file, cur_path):
    print(cur_path)
    excel_reader = ExcelReader(xlsx_file)
    rows_count = excel_reader.rows()
    for i in range(rows_count):
        id = excel_reader.get_content(i, "身份证号")
        docx_writer = WordWriter('./special_building_template.docx', f"{cur_path}/{id}.docx")
        docx_writer.set_id(id)
        name = excel_reader.get_content(i, "姓名")
        docx_writer.set_name(name)
        sex = excel_reader.get_content(i, "性别")
        docx_writer.set_sex(sex)
        birth = excel_reader.get_content(i, "出生年月")
        docx_writer.set_birth(birth)
        work_code = excel_reader.get_content(i, "操作类别（工种）")
        docx_writer.set_work_code(work_code)
        address = excel_reader.get_content(i, "通讯地址")
        docx_writer.set_address(address)
        company = excel_reader.get_content(i, "单位")
        docx_writer.set_company(company)
        phone = excel_reader.get_content(i, "联系电话")
        docx_writer.set_phone(phone)

        print(f"img path: {cur_path}/{id}.jpeg")
        if os.path.exists(f"{cur_path}/{id}.jpeg") is True:
            docx_writer.set_photo(f"{cur_path}/{id}.jpeg")

        docx_writer.save()


def automate_excel_docx(upload_zip_file):
    with tempfile.TemporaryDirectory() as tmp_dir_name:
        with zipfile.ZipFile(upload_zip_file, 'r') as zip_ref:
            zip_ref.extractall(tmp_dir_name)

        zip_dir = Path(upload_zip_file).stem
        cur_path = f"{tmp_dir_name}/{zip_dir}"
        xlsx_files = glob.glob(os.path.join(cur_path, "*.xlsx"))
        for xlsx_file in xlsx_files:
            extract_xlsx_docx(xlsx_file, cur_path)

        with zipfile.ZipFile('./complete.zip', 'w') as f:
            for file in glob.glob(os.path.join(cur_path, "*.docx")):
                f.write(file)


automate_excel_docx("./report.zip")



