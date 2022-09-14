from report_constructor import *
from get_external_info import *
from docx import Document


def generate_file(path_with_pictures, report_path):

    # path = 'C:/Users/PC/Desktop/Work/'  # legacy for save time
    # path_with_pictures = path + '/Pictures'
    # path_with_pictures = select_path()
    report_doc = Document()
    set_margin(report_doc)
    all_pictures = get_pictures_list(path_with_pictures)
    all_floors = get_floor_list(all_pictures)
    # report_path = select_path()

    # report_path = path

    if save_report(report_doc, report_path) != 1:
        add_footer_in_doc(report_doc)
        add_header_in_doc(report_doc)
        add_first_page_in_doc(report_doc)
        absolut_picture_number_in_file = 1
        for current_floor in all_floors:
            current_floor_pictures_list = get_list_of_floor_pictures(current_floor, all_pictures)
            absolut_picture_number_in_file = add_floor_in_report(path_with_pictures, report_doc, current_floor,
                                                                 current_floor_pictures_list,
                                                                 absolut_picture_number_in_file)
        save_report(report_doc, report_path)
        print("SUCCESS")
    else:
        print("CLOSE FILE")
