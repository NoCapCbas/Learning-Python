import sys
import pytest
import requests

# script used to test the TDM Api

class TestTdmApi:
	# test account
	USERNAME = 'test'
	PASSWORD = 'Carryon2020'

	def test_username_password_auth(self):
		url = f'https://www1.tdmlogin.com/tdm/api/api.asp?username={self.USERNAME}&password={self.PASSWORD}&flow=B&reporter=US&partners=AFGHAN&periodBegin=201701&periodEnd=201701'

		r = requests.get(url)

		print(f'\nRequest Status Reason: {r.reason}')
		print(f'Request Elapsed Total Seconds: {r.elapsed.total_seconds()}s')
		assert r.status_code == 200, f'Status Code: {r.status_code}, Expected Status Code: 200'

		assert len(r.text) > 0, f'Length of Response Data: {len(r.text)}, Expected Length of Response Data > 0'

	def test_api_key_auth(self):
		pass

	def test_complex_case_4(self):
		# example 4 in API_Specification_Stat doc
		url = f'https://www1.tdmlogin.com/tdm/api/api.asp?username={self.USERNAME}&password={self.PASSWORD}&flow=E&reporter=US&partners=CANADA,SPAIN&periodBegin=201701&periodEnd=209912&currency=EUR&hsCode=01,02,03,04,05&levelDetail=6&includeUnits=BOTH'
		
		r = requests.get(url)

		print(f'\nRequest Status Reason: {r.reason}')
		print(f'Request Elapsed Total Seconds: {r.elapsed.total_seconds()}s')
		assert r.status_code == 200, f'Status Code: {r.status_code}, Expected Status Code: 200'

		assert len(r.text) > 0, f'Length of Response Data: {len(r.text)}, Expected Length of Response Data > 0'

	def test_complex_case_5(self):	
		# example 5 in API_Specification_Stat doc
		url = f'https://www1.tdmlogin.com/tdm/api/api.asp?username={self.USERNAME}&password={self.PASSWORD}&flow=E&reporter=US&partners=CANADA,SPAIN&periodBegin=201701&periodEnd=209912&currency=EUR&hsCode=01,02,03,04,05&levelDetail=6&includeUnits=BOTH&isoCountryCode=BOTH&includeDesc=Y&lang=ES'
		
		r = requests.get(url)

		print(f'\nRequest Status Reason: {r.reason}')
		print(f'Request Elapsed Total Seconds: {r.elapsed.total_seconds()}s')
		assert r.status_code == 200, f'Status Code: {r.status_code}, Expected Status Code: 200'

		assert len(r.text) > 0, f'Length of Response Data: {len(r.text)}, Expected Length of Response Data > 0'

if __name__ == '__main__':
	sys.exit(pytest.main(['-s', '-v', 'TestTdmApi.py']))

