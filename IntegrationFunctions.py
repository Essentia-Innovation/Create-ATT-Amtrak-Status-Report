from requests.auth import HTTPBasicAuth
import requests
import json


class BasicIntegration:
    """Main class for interacting with OneVizion API and managing espeed data"""

    def __init__(self, urlOneVizion="", loginOneVizion="", passOneVizion=""):
        """Initialize ToeSpeed with necessary parameters for API"""
        # Assign variables
        self.url_onevizion = urlOneVizion
        self.login_onevizion = loginOneVizion
        self.pass_onevizion = passOneVizion
        self.auth_onevizion = HTTPBasicAuth(self.login_onevizion, self.pass_onevizion)
        # Define headers for API calls
        self.headers = {'Content-type': 'application/json', 'Content-Encoding': 'utf-8'}


    def update_trackor(self, child_trackor_id, child_dict : dict):
        """Update a single trackor using OneVizion API"""
        url = 'https://{url_onevizion}/api/v3/trackors/{child_trackor_id}'.format(url_onevizion=self.url_onevizion,
                                                                                  child_trackor_id=child_trackor_id)
        answer = requests.put(url, data=json.dumps(child_dict), headers=self.headers, auth=self.auth_onevizion)
        if answer.ok:
            return answer.json()
        else:
            raise Exception(answer.text)

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

    def create_trackor(self, child, child_field : dict, parent, parent_field : dict):
        """Create a trackor using OneVizion API"""
        url = 'https://{url_onevizion}/api/v3/trackor_types/{child_trackor_id}/trackors'.format(url_onevizion=self.url_onevizion,
                                                                                  child_trackor_id=child)
        data = {'fields': child_field, 'parents': [{'trackor_type': parent, 'filter': parent_field}]}
        answer = requests.post(url, data=json.dumps(data), headers=self.headers, auth=self.auth_onevizion)
        if answer.ok:
            return answer.json()
        else:
            raise Exception(answer.text)

    def create_trackor_noparent(self, child, child_field : dict,):
        """Create a trackor using OneVizion API"""
        url = 'https://{url_onevizion}/api/v3/trackor_types/{child_trackor_id}/trackors'.format(url_onevizion=self.url_onevizion,
                                                                                  child_trackor_id=child)
        data = {'fields': child_field}
        answer = requests.post(url, data=json.dumps(data), headers=self.headers, auth=self.auth_onevizion)
        if answer.ok:
            return answer.json()
        else:
            raise Exception(answer.text)

    def delete_trackor(self, trackor, track_id,):
        """Delete a trackor using OneVizion API"""
        url = 'https://{url_onevizion}/api/v3/trackor_types/{trackor_type}/trackors?trackor_id={trackor_ID}'.format(url_onevizion=self.url_onevizion,
                                                                                 trackor_type=trackor, trackor_ID=track_id)
        answer = requests.delete(url, data=json.dumps(trackor), headers=self.headers, auth=self.auth_onevizion)
        if answer.ok:
            return ("Trackor Deleted")
        else:
            raise Exception(answer.text)
