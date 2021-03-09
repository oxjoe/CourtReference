import re
import time
from datetime import datetime

import requests
import xlsxwriter
from bs4 import BeautifulSoup


def main():
    main_url = 'https://www.courtreference.com'
    now = datetime.now().strftime('%m_%d_%Y_%H_%M_%S')
    output_filename = f'demo_{now}.xlsx'
    main_wksht_name = 'MAIN_SHEET'

    page = requests.get(main_url)
    soup = BeautifulSoup(page.content, 'html.parser')

    list_of_states_html = soup.find(id='homeStateList')
    states_to_iterate = list_of_states_html.find_all('a')
    # states_to_iterate = list_of_states_html.find_all('a', title='New Hampshire Court Records and Guide')
    # states_to_iterate = list_of_states_html.find_all('a', title='New Jersey Court Records and Guide')
    states_left = len(states_to_iterate)
    table_headers_list = ['District Name', 'Court Address', 'Court Contact Information', 'County URL',
                          'Direct Court URL']
    number_of_columns_for_each_county = len(table_headers_list)
    workbook = xlsxwriter.Workbook(output_filename)
    worksheet = workbook.add_worksheet(main_wksht_name)

    # Some style settings
    bold = workbook.add_format({'bold': True})
    red = workbook.add_format({'font_color': 'red'})

    # Create the first sheet that links to all the other state sheets (aka the first sheet)
    for count, state in enumerate(states_to_iterate):
        sheet_name = state.text
        worksheet.set_column('A:A', 40)
        worksheet.write_url(f'A{count + 1}', f"internal:'{sheet_name}'!A1", string=sheet_name)

    def get_list_of_counties(state_url):
        page = requests.get(state_url)
        soup = BeautifulSoup(page.content, 'html.parser')
        county_list = soup.find(class_='dropdown-menu').find_all('a')
        return county_list

    def get_county_elems(county_url):
        page = requests.get(county_url)
        soup = BeautifulSoup(page.content, 'html.parser')
        return soup

    def write_subheaders(i, j_dict, lst):
        for val in j_dict.values():
            for item in lst:
                worksheet.write(i, val, item, red)
                val += 1

    def write_data(i, j, lst, display_name):
        for index, value in enumerate(lst):
            if index == 3:
                worksheet.write_url(i, j, value, string=display_name)
            else:
                worksheet.write(i, j, value)
            j += 1

    # Create a worksheet for each state
    for state in states_to_iterate:
        print(f"States Left: {states_left}")
        states_left -= 1
        state_name = state.text
        attributes = state.attrs
        worksheet = workbook.add_worksheet(state_name)
        direct_state_court_url = attributes.get('href')
        worksheet.write_url(f'A1', f"internal:'{main_wksht_name}'!A1", string=f'GO BACK TO {main_wksht_name}')
        state_url = main_url + direct_state_court_url
        worksheet.write('B1', state_url)
        county_elems = get_list_of_counties(state_url)

        #######################
        # Regex for matching against stuff like Circuit Courts in Autauga County
        pattern = r'\s(in|IN).*'
        # Tracking column number
        court_type_header_dict = {}
        # Tracking row number
        row_dict = {}
        # County data will start at excel row 4 of each excel sheet (rows are zero indexed)
        starting_row = 3
        # Log for user to see how many counties + states are left
        progress = len(county_elems)
        #######################

        for county in county_elems:
            county_name = county.text
            attributes = county.attrs
            progress -= 1
            print(f"On County: {county_name} of {state_name} with {progress} counties left...")

            final_direct_court_url = attributes.get('href')
            final_county_url = main_url + final_direct_court_url

            county_main_html = get_county_elems(final_county_url)

            court_headers_elems = county_main_html.find_all('h3', class_='titl')
            num_court_headers_for_current_county = len(court_headers_elems)

            # Nice county name to be displayed in place of the URL in excel
            init_element = county_main_html.find('h1')
            if init_element is None:
                print("Above county doesn't have a heading so it's probably blank...")
            else:
                header = county_main_html.find('h1').text.strip()
                okay_pattern = r'\s(County|Borough).+'
                # Autauga County Alabama Court Directory
                # Fairbanks North Star Borough Alaska Court Directory
                # ABC World View County California Court Directory
                bolded_county = re.split(okay_pattern, header)[0]
                # print(header)
                # print(bolded_county)
                if 'County' in header.split() and 'Borough' not in header.split():
                    county_display_name = bolded_county + ' County'
                elif 'Borough' in header.split() and 'County' not in header.split():
                    county_display_name = bolded_county + ' Borough'
                else:
                    county_display_name = "FIGURE_IT_OUT_YOURSELF"

            for index_x, value_x in enumerate(court_headers_elems):
                header_text = value_x.text.upper()
                # print(header_text)
                final_court_type = re.split(pattern, header_text)[0]

                # If not in dict then add it, else that means it is new and must be added with a bigger value
                if final_court_type not in court_type_header_dict and not court_type_header_dict:
                    # Adding all the court types for the first time
                    court_type_header_dict[final_court_type] = index_x * number_of_columns_for_each_county
                    row_dict[final_court_type] = starting_row
                    # Write court type header
                    worksheet.write(1, court_type_header_dict.get(final_court_type), final_court_type, bold)
                    # Write subheaders
                    write_subheaders(2, court_type_header_dict, table_headers_list)
                elif final_court_type not in court_type_header_dict:
                    # Adding a brand new court type
                    new_index = max(court_type_header_dict.values()) + number_of_columns_for_each_county
                    court_type_header_dict[final_court_type] = new_index
                    row_dict[final_court_type] = starting_row
                    # Write court type header
                    worksheet.write(1, court_type_header_dict.get(final_court_type), final_court_type, bold)
                    # Write subheaders
                    write_subheaders(2, court_type_header_dict, table_headers_list)
                # print(result)
            # print(court_type_header_dict)
            # print(row_dict)

            court_type_group_elems = county_main_html.find_all('div', class_='court-type-group')
            if len(court_headers_elems) != len(court_type_group_elems):
                raise Exception(
                    f'# of Court Groups = [{len(court_type_group_elems)}] are not matching up with the # of Green '
                    f'Headers = [{len(court_headers_elems)}]')

            count = 0
            # print(num_court_headers_for_current_county)
            while count < num_court_headers_for_current_county:
                green_header_elem = county_main_html.find_all('h3', class_='titl')[count]
                init_header_text = green_header_elem.text.upper()
                final_header_text = re.split(pattern, init_header_text)[0]
                if final_header_text in court_type_header_dict:
                    target_column = court_type_header_dict.get(final_header_text)
                else:
                    raise Exception("Something messed up. Court Type headers aren't matching so "
                                    "unable to determine target_column.")

                # Goto court group
                court_type_group_el = green_header_elem.find_next_sibling("div")
                # print(court_type_group_el)
                article_el = court_type_group_el.find_all('article', class_='county-result-entry')
                # print(article_el)
                for article in article_el:
                    first_text = article.find('a', class_='court-info')
                    attributes = first_text.attrs
                    district_name = attributes.get('title').strip()
                    final_direct_court_url = main_url + attributes.get('href')

                    init_address = article.find('div', property='address')
                    if init_address is None:
                        final_address = 'No Address Found'
                    else:
                        final_address = " ".join(init_address.text.split())
                    # print(final_address)

                    init_phone = article.find('span', property='telephone')
                    if init_phone is None:
                        final_phone = 'Phone: No Phone Found'
                    else:
                        phones = article.find_all('span', property='telephone')
                        phone_list = []
                        for phone in phones:
                            # a = soup.find('span', property='telephone')
                            # a
                            # <span property="telephone">574-235-9794</span>
                            # a.parent
                            # <span>Phone: <span property="telephone">574-235-9794</span> (Small Claims)</span>
                            # a.parent.text
                            # 'Phone: 574-235-9794 (Small Claims)'
                            init_phone = phone.parent.text.strip()
                            phone_list.append(init_phone)
                        final_phone = '\n'.join(phone_list)
                        # final_phone = init_phone.text.strip()
                    # print(final_phone)

                    init_fax = article.find('span', property='faxNumber')
                    if init_fax is None:
                        final_fax = 'No Fax Found'
                    else:
                        final_fax = init_fax.text.strip()
                    # print(final_fax)

                    final_data = [district_name,
                                  final_address,
                                  f"{final_phone}" + '\n' + f"Fax: {final_fax}",
                                  final_county_url,
                                  final_direct_court_url]
                    row = row_dict.get(final_header_text)
                    col = target_column
                    write_data(row, col, final_data, county_display_name)
                    row_dict[final_header_text] += 1

                count += 1

    workbook.close()


if __name__ == "__main__":
    tic = time.perf_counter()
    main()
    toc = time.perf_counter()
    print(f"Took about {toc - tic:0.4f} seconds or about {(toc - tic) / 60:0.1f} minutes")
