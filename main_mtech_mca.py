import requests, re
from urllib.request import urlopen, Request
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import pandas as pd
from PyQt5.QtCore import QThread, pyqtSignal


class PG(QThread):
    cbcs_update_progress = pyqtSignal(float)

    def __init__(self, parent_url, start_usn, end_usn, file_path=None, parent=None):
        super(PG, self).__init__(parent)
        self.parent_url = parent_url
        self.file_path = file_path
        self.start_usn = start_usn
        self.end_usn = end_usn
        self.usn_base = self.start_usn[:-2]
        self.child_url = 'http://results.vtu.ac.in/vitaviresultcbcs2018/resultpage.php'
        self.regexp = '^\d{2}[A-Z]{2}'
        self.final_df = pd.DataFrame()
        self.sess = requests.session()
        self.token = None

    def get_token(self):
        req = Request(self.parent_url, headers={'User-Agent': 'Mozilla/5.0'})
        html = urlopen(req).read()
        temp_soup = BeautifulSoup(html, 'html5lib')
        self.token = temp_soup.find("input", {"id": "tokenid"})['value']
        return self.token

    def run(self):
        token = self.get_token()

        for usn in range(int(self.start_usn[8:]), int(self.end_usn[8:]) + 1):

            if len(str(usn)) < 2:
                usn_str = '0' * (2 - len(str(usn))) + str(usn)
            else:
                usn_str = str(usn)

            usn_full = self.usn_base + usn_str
            print(usn_full)
            progress = (usn - int(self.start_usn[8:])) / (int(self.end_usn[8:]) - int(self.start_usn[8:])) * 100
            self.cbcs_update_progress.emit(progress)

            result_data = dict(lns=usn_full, token=token, current_url=self.parent_url)
            page = self.sess.post(self.child_url, data=result_data).content
            soup = BeautifulSoup(page, 'html5lib')

            try:
                name = soup.find("td", {"style": "padding-left:15px"}).text.strip(': ')
                all_details = soup.findAll("div", {"class": "divTableCell"})[6:-5]
            except:
                continue

            all_details = [div.text for div in all_details]

            if len(all_details) < 1:
                continue

            sub_codes = [code for code in all_details if re.match(self.regexp, code)]

            marks_dict = dict()
            all_subs_df = pd.DataFrame()
            for sub_code in sub_codes:
                index = all_details.index(sub_code)

                try:
                    internal_marks = int(all_details[index + 2])
                    external_marks = int(all_details[index + 3])
                    subject_total_marks = int(all_details[index + 4])
                except:
                    internal_marks = all_details[index + 2]
                    external_marks = all_details[index + 3]
                    subject_total_marks = all_details[index + 4]

                pass_fail = all_details[index + 5]

                marks_dict[sub_code] = [internal_marks, external_marks, subject_total_marks, pass_fail]

                sub_df = pd.DataFrame(marks_dict[sub_code],
                                      index=['1.Internal', '2.External', '3.Subject Total', '4.Pass/Fail'])
                temp_df = pd.concat([sub_df], keys={sub_code: sub_df})
                temp_df = temp_df.transpose()
                all_subs_df = pd.concat([all_subs_df, temp_df], axis=1)

            all_subs_df = all_subs_df.rename_axis({0: name}, axis='index')
            df = pd.concat([all_subs_df], keys={usn_full: all_subs_df})
            self.final_df = pd.concat([self.final_df, df])

        self.write_excel()

    def write_excel(self):
        if len(self.file_path) < 1:
            self.file_path = 'CBCS_Result.xlsx'
        elif self.file_path[-5:] == '.xlsx':
            self.file_path = self.file_path
        else:
            self.file_path = self.file_path + '.xlsx'

        try:
            book = load_workbook(self.file_path)
            writer = pd.ExcelWriter(self.file_path, engine='openpyxl')
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            self.final_df.to_excel(writer, self.start_usn + '_' + self.end_usn)
            writer.save()
            self.sess.close()
        except:
            writer = pd.ExcelWriter(self.file_path, engine='openpyxl')
            self.final_df.to_excel(writer, self.start_usn + '_' + self.end_usn)
            writer.save()
            self.sess.close()


if __name__ == '__main__':
    url = 'http://results.vtu.ac.in/vitaviresultcbcs2018/index.php'
    test_obj = PG(url, '1RN16LVS01', '1RN16LVS16', 'MCA_MTECH_Result.xlsx')
    test_obj.run()
