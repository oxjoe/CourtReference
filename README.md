# CourtReference

- See the excel file(s) for an example output.

- Some courts have the court name as part of the Address field, it just grabs the entire address field so it will be pulled
  in too.

- ~51 states takes about 36 mins

- Contact Info LOOKS like Phone and Fax are NOT a space apart, but they are actually new line apart.

- User has to be careful and inspect the County URL column b/c of:

  ```python
  if 'County' in header.split() and 'Borough' not in header.split():
    county_display_name = bolded_county + ' County'
  elif 'Borough' in header.split() and 'County' not in header.split():
    county_display_name = bolded_county + ' Borough'
  else:
    county_display_name = "FIGURE_IT_OUT_YOURSELF"
  ```
  
  

