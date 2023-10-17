from openpyxl import load_workbook
from glob import glob
import pandas
import json
from requests.auth import HTTPBasicAuth
import requests


verifyColumns = False


class RFSUpdate:
    def __init__(self):
        with open('settings.json', "rb") as PFile:
            password_data = json.loads(PFile.read().decode('utf-8'))

            """Initialize ToeSpeed with necessary parameters for API"""
            # Assign variables

            self.url_onevizion = password_data["urlOneVizion"]
            self.login_onevizion = password_data["loginOneVizion"]
            self.pass_onevizion = password_data["passOneVizion"]
            self.auth_onevizion = HTTPBasicAuth(self.login_onevizion, self.pass_onevizion)
            # Define headers for API calls
            self.headers = {'Content-type': 'application/json', 'Content-Encoding': 'utf-8'}


        # Generate list of client dictionaries from eSpeed to compare against spreadsheet
        rfsdict = self.search_trackors("ENTProject", ["XITOR_KEY", "ENTPR_RFS_PHASE", "ENTPR_RFS_ACTIVE_TASK",
                                                             "ENTPR_NOTES_HISTORY"],
                                              'is_not_null(XITOR_KEY)')

        file = glob("*.xlsx")[0]
        wb = load_workbook(file)
        ws = wb.worksheets[0]
        sheet_df = pandas.DataFrame(ws.values)
        new_header = sheet_df.iloc[0]  # grab the first row for the header
        sheet_df = sheet_df[1:]  # take the data less the header row
        sheet_df.columns = new_header  # set the header row as the df header


        if new_header[0] == 'RFS Number' and new_header[8] == 'Essentia-RFS Phase Update' and new_header[
            10] == 'Essentia-RFS Active  Task' and new_header[11] == 'Essentia ETA-Update':
            verifyColumns = True

        if verifyColumns:
            for rowNum in range(2, ws.max_row + 1):
                if ws.cell(rowNum, 1).value is not None:
                    for i in rfsdict:
                        if ws.cell(rowNum, 1).value == i["TRACKOR_KEY"]:
                            print("MATCH: ")
                            print("Espeed Data: ", i["TRACKOR_KEY"], "Phase: ", i["ENTPR_RFS_PHASE"], "Active Task: ",
                                  i["ENTPR_RFS_ACTIVE_TASK"], "Notes:", i["ENTPR_NOTES_HISTORY"])
                            print("Excel Data: ", ws.cell(rowNum, 1).value, "Phase: ", ws.cell(rowNum, 9).value,
                                  "Active Task: ", ws.cell(rowNum, 11).value, "Notes: ", ws.cell(rowNum, 12).value)
                            ws.cell(rowNum, 12).value = i["ENTPR_NOTES_HISTORY"]
                            if ws.cell(rowNum, 8).value != i["ENTPR_RFS_PHASE"] or ws.cell(rowNum, 10).value != i[
                                "ENTPR_RFS_ACTIVE_TASK"]:
                                ws.cell(rowNum, 9).value = i["ENTPR_RFS_PHASE"]
                                ws.cell(rowNum, 11).value = i["ENTPR_RFS_ACTIVE_TASK"]
                            print("New Excel Data: ", ws.cell(rowNum, 1).value, "Phase: ", ws.cell(rowNum, 9).value,
                                  "Active Task: ", ws.cell(rowNum, 11).value, "Notes: ", ws.cell(rowNum, 12).value)
                            print("")

            wb.save('output.xlsx')
        else:
            print("Column names do not match expected columns.")

    def search_trackors(self, trackor_type: str, fields: list, filter: str):
            """Search for trackors through a filter using OneVizion API.

            Parameters
            ----------
            trackor_type : str
                Trackor Type of the Trackor for which data should be retrieved (e.g., BILL_OF_MATERIALS).
                Use the Trackor Type name, not the Trackor Type label.

            fields: list
                Comma separated list of Field names to return in response.
                Note: If Field doesn't exists HTTP 404 (Not found) will be returned

            filter: str
                Supported filter operators:
                equal(cf_name, value) or =(cf_name, value)
                greater(cf_name, value) or >(cf_name, value)
                less(cf_name, value) or <(cf_name, value)
                gt_today(cf_name, value = +0) or >=Today(cf_name, value = +0) [use + or - in front of value without spaces, if not specified +0 will be used]
                lt_today(cf_name, value = +0) or <=Today(cf_name, value = +0) [use + or - in front of value without spaces, if not specified +0 will be used]
                within(cf_name, value1, value2)
                not_equal(cf_name, value) or <>(cf_name, value)
                null(cf_name) or is_null(cf_name)
                is_not_null(cf_name)
                outer_equal(cf_name, value)
                outer_not_equal(cf_name, value)
                less_or_equal(cf_name, value) or <=(cf_name, value)
                greater_or_equal(cf_name, value) or >=(cf_name, value)
                this_week(cf_name, value = +0) [use + or - in front of value without spaces, if not specified +0 will be used]
                this_month(cf_name, value = +0) [use + or - in front of value without spaces, if not specified +0 will be used]
                this_quarter(cf_name, value = +0) [use + or - in front of value without spaces, if not specified +0 will be used] (equal to This FQ)
                this_year(cf_name, value = +0) [use + or - in front of value without spaces, if not specified +0 will be used] (equal to This FY)
                this_week_to_date(cf_name)
                this_month_to_date(cf_name)
                this_quarter_to_date(cf_name) (equal to This FQ to Dt)
                this_year_to_date(cf_name) (equal to This FY to Dt)
                field_equal(cf1_name, cf2_name) or =F(cf1_name, cf2_name)
                field_not_equal(cf1_name, cf2_name) or <>F(cf1_name, cf2_name)
                field_less(cf1_name, cf2_name) or <F(cf1_name, cf2_name)
                field_greater(cf1_name, cf2_name) or >F(cf1_name, cf2_name)
                field_less_or_equal(cf1_name, cf2_name) or <=F(cf1_name, cf2_name)
                field_greater_or_equal(cf1_name, cf2_name) or >=F(cf1_name, cf2_name)
                new(cf_name) or is_new(cf_name)
                not_new(cf_name) or is_not_new(cf_name)
                equal_myself(cf_name) or =Myself(cf_name)
                not_equal_myself(cf_name) or <>Myself(cf_name)

            Raises
            ------
            Exception
                If the response from onevizion is not ok
            """

            url = 'https://{url_onevizion}/api/v3/trackor_types/{trackor_type}/trackors/search?fields={fields}'.format(
                url_onevizion=self.url_onevizion,
                trackor_type=trackor_type, fields=",".join(fields))
            answer = requests.post(url, data=filter, headers=self.headers, auth=self.auth_onevizion)
            if answer.ok:
                return answer.json()
            else:
                raise Exception(answer.text)

if __name__ == "__main__":
        RFSUpdate()
