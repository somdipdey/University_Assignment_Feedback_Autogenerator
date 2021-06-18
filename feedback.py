# Reading an excel file using Python
import xlrd
from reportlab.lib.colors import blue
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import inch
from reportlab.pdfgen.canvas import Canvas
# To generate pdf using python refer to https://realpython.com/creating-modifying-pdf/

# Give the file_location of the excel file
file_loc = ("CEXXX-Assignment1-feedback.xls")

# To open Workbook
wb = xlrd.open_workbook(file_loc)
sheet = wb.sheet_by_index(0) #choose sheet number starting from 0

# XXXCHANGE HERE
# choose the default row value where the marks start
row_num = 1
total_rows = 5 # provide the last row number which has records

while row_num < total_rows:
    print("Processing feedback for " + str(int(sheet.cell_value(row_num, 0))))
    # XXXCHANGE HERE
    # change the following parameters for your set of questions
    registration_id = str(int(sheet.cell_value(row_num, 0)))
    q1_col_record = str(int(sheet.cell_value(row_num, 1))) + "%"
    q2_col_record = str(int(sheet.cell_value(row_num, 2))) + "%"
    q3_col_record = str(int(sheet.cell_value(row_num, 3))) + "%"
    q4_col_record = str(int(sheet.cell_value(row_num, 4))) + "%"
    q5_col_record = str(int(sheet.cell_value(row_num, 5))) + "%"
    q6_col_record = str(int(sheet.cell_value(row_num, 6))) + "%"
    q7_col_record = str(int(sheet.cell_value(row_num, 7))) + "%"
    total_record = str(int(sheet.cell_value(row_num, 8))) + "%"
    feedback_comment = sheet.cell_value(row_num, 9)

    # XXXCHANGE HERE
    # change the following parameters for your set of questions
    q1 = "Question 1 (25%): "
    q2 = "Question 2 (25%): "
    q3 = "Question 3 (10%): "
    q4 = "Question 4 (10%): "
    q5 = "Question 5 (10%): "
    q6 = "Question 6 (10%): "
    q7 = "Question 7 (10%): "
    total_marks = "Total (100%): "

    # XXXCHANGE HERE
    # specify pdf name pattern
    pdf_file_name = "feedback/" + registration_id + "_feedback.pdf"
    canvas = Canvas(pdf_file_name, pagesize=LETTER)

    # Set font to Times New Roman with 12-point size
    canvas.setFont("Times-Roman", 15)

    # Draw blue text one inch from the left and ten
    # inches from the bottom
    canvas.setFillColor(blue)

    # marks for each question is set here ->
    question_base_line_gap = 10
    canvas.drawString(1 * inch, question_base_line_gap * inch, "Registration ID: " + registration_id)
    canvas.drawString(1 * inch, (question_base_line_gap - 0.9) * inch, q1 + q1_col_record)
    canvas.drawString(1 * inch, (question_base_line_gap - 1.2) * inch, q2 + q2_col_record)
    canvas.drawString(1 * inch, (question_base_line_gap - 1.5) * inch, q3 + q3_col_record)
    canvas.drawString(1 * inch, (question_base_line_gap - 1.8) * inch, q4 + q4_col_record)
    canvas.drawString(1 * inch, (question_base_line_gap - 2.1) * inch, q5 + q5_col_record)
    canvas.drawString(1 * inch, (question_base_line_gap - 2.4) * inch, q6 + q6_col_record)
    canvas.drawString(1 * inch, (question_base_line_gap - 2.7) * inch, q7 + q7_col_record)
    canvas.drawString(1 * inch, (question_base_line_gap - 3.0) * inch, total_marks + total_record)

    # feedback comments printing in the pdf starts here ->
    start_character = 0
    end_character = 75
    comment_base_line_gap = 5.0
    line_space = 0
    length_of_comment = len(feedback_comment)
    canvas.drawString(1 * inch, (comment_base_line_gap + 0.6) * inch, "Feedback: ")
    if (length_of_comment > end_character):
        while (end_character < length_of_comment):
            canvas.drawString(1 * inch, (comment_base_line_gap - line_space) * inch, feedback_comment[start_character:end_character] + "-")
            start_character += 75
            end_character += 75
            line_space += 0.3
        end_character -= 75
        canvas.drawString(1 * inch, (comment_base_line_gap - line_space) * inch, feedback_comment[end_character:len(feedback_comment)])
    else:
        canvas.drawString(1 * inch, comment_base_line_gap * inch, feedback_comment)

    # Save the PDF file
    canvas.save()

    #updated the row number
    row_num += 1
