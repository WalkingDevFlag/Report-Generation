from docx import Document
import pandas as pd
def generate_report(template_path, output_path, data):
    doc = Document(template_path)

    for paragraph in doc.paragraphs:
        for key, value in data.items():
            if key in paragraph.text:
                for run in paragraph.runs:
                    run.text = run.text.replace(key, value)
    
    doc.save(output_path)


def generate_report_csv(csv_path, template_path):
    df = pd.read_csv('contacts.csv')
    for idx, row in df.iterrows():
        data = {
            '[Salutation]': row['Salutations'],
            '[First Name]': row['First Name'],
            '[Last Name]': row['Last Name'],
            '[Last Contacted]': row['Last Contacted'],
            '[Company Name]':row['Company Name'] 
        }

        output_path = f'report_{idx + 1}.docx'
        generate_report(template_path, output_path, data)


if __name__ == '__main__':
    generate_report_csv('contacts.csv', 'template2.docx')
    # data = {
    #     '[Salutation]': 'Mr.',
    #     '[First Name]': 'Amey',
    #     '[Last Name]': 'Darokar',
    #     '[Last Contacted]': '7th Janurary 2025',
    #     '[Company Name]': 'ASG'
    # }

    # template_path = 'template2.docx'
    # output_path = 'Ouput.docx'

    # generate_report(template_path, output_path, data)