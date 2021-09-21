import glob
import os
import shutil
from PyPDF2 import PdfFileMerger
from os.path import exists
import tkinter as tk


def previous_designs(entry, response_label, window):

    job_number = entry.get()
    response_label.config(text='Program is processing...')
    window.update()

    # creates temp test file if exists
    if not os.path.exists(r"C:/Users/ltaylor/Desktop/test123"):
        os.makedirs(r"C:/Users/ltaylor/Desktop/test123")

    # returns a list of names in list files.
    tenant_search = r'G:/' + job_number.replace(job_number[-3:], '000' + '*')
    files = glob.glob(tenant_search, recursive=False)
    if len(files) == 1:
        building_folder = files[0]

        # finds jobs inside job folder
        job_folders = glob.glob(building_folder + '/*')

        # gets list of previous jobs inside job folder, filtering out other files
        previous_jobs_list = []
        for folder in job_folders:
            folder_name = folder.split('\\')[-1]
            if folder_name[0].isnumeric() and folder_name[1].isnumeric() and folder_name[2].isnumeric():
                previous_jobs_list.append(folder)

        # gets floor of current job
        job_floor = 'unassigned'
        for previous_job in previous_jobs_list:
            if previous_job.split('\\')[-1][:3] == job_number[-3:]:
                job_floor = previous_job[-2:]

        # gets existing plans folder inside job folder
        try:
            existing_plans_folder = glob.glob(building_folder + '/' + job_number[-3:] + '*/existing plans*')[0] + '/'
        except:
            response_label.config(text='Job number is invalid.')
            return

        # checks if existing plans folder already has the correct files, exits function if so
        existing_plans_content = glob.glob(existing_plans_folder + '/*')
        for file in existing_plans_content:
            if file.split('\\')[-1] == 'all_previous_designs.pdf':
                response_label.config(text='Previous jobs have already been compiled.')
                return

        # gets most recent mechanical pdf for each previous job and copies it to existing plans folder
        num = 1     # assigns number to each pdf that is being copied, used when naming copied pdf
        for previous_job in previous_jobs_list:
            job_issuances = glob.glob(previous_job + "/Pdf's/*")
            if len(job_issuances) > 0:
                valid_issuances = []
                for issuance in job_issuances:
                    if issuance.split('\\')[-1][:3].isnumeric():    # adds file to issuance list if first 4 are numbers
                        valid_issuances.append(issuance)
                valid_issuances.reverse()   # reverses order of list, make most recent first
                pdf = 'unassigned'
                for issuance2 in valid_issuances:
                    for file in glob.glob(issuance2 + '/*', recursive=False):   # checks if issuance has mech file
                        if 'M-1' in file.split('/')[-1] or 'MECH' in file.split('/')[-1]:
                            pdf = file
                            shutil.copy(pdf, r"C:/Users/ltaylor/Desktop/test123/"
                                        + pdf.split('\\')[-1].removesuffix('.pdf') + '-' + str(num) + '.pdf')
                            num = num + 1
                            break
                    if pdf != 'unassigned':
                        #print(pdf)
                        break
    else:
        print('Returned too many directories.')
    pdf_combine(job_number, job_floor, existing_plans_folder)
    response_label.config(text='Previous jobs have been compiled into Existing Plans folder.')


# combines previous job pdf's into one. Compiles same floor first if there are enough, then the rest.
def pdf_combine(job_number, job_floor, existing_plans_folder):
    all_pdfs = glob.glob(r"C:/Users/ltaylor/Desktop/test123/*", recursive=False)

    # makes list of files for current job floor
    current_floor_pdfs = []
    for file in all_pdfs:
        if file.split('\\')[-1][:2] == job_floor:
            current_floor_pdfs.append(file)

    # combines pdfs of the same floor
    if len(current_floor_pdfs) > 0:
        same_floor_merger = PdfFileMerger(strict=False)
        for pdf in current_floor_pdfs:
            same_floor_merger.append(pdf)
        same_floor_merger.write(r"C:/Users/ltaylor/Desktop/test123/same_floor_designs.pdf")
        same_floor_merger.close()
        source_file1 = r"C:/Users/ltaylor/Desktop/test123/same_floor_designs.pdf"
        dest_file1 = existing_plans_folder + 'same_floor_designs.pdf'
        shutil.copy(source_file1 , dest_file1)

    # combines pdfs of all the floors
    all_floors_merger = PdfFileMerger(strict=False)
    for pdf in all_pdfs:
        all_floors_merger.append(pdf)
    all_floors_merger.write(r"C:/Users/ltaylor/Desktop/test123/all_previous_designs.pdf")
    all_floors_merger.close()
    source_file2 = r"C:/Users/ltaylor/Desktop/test123/all_previous_designs.pdf"
    dest_file2 = existing_plans_folder + 'all_previous_designs.pdf'
    shutil.copy(source_file2, dest_file2)

    # removes all excess pdf's
    for file in all_pdfs:
        os.remove(file)
    if exists(r"C:/Users/ltaylor/Desktop/test123/same_floor_designs.pdf"):
        os.remove(r"C:/Users/ltaylor/Desktop/test123/same_floor_designs.pdf")
    os.remove(r"C:/Users/ltaylor/Desktop/test123/all_previous_designs.pdf")
    os.rmdir(r"C:/Users/ltaylor/Desktop/test123")



# opens job folder
def open_job_folder(entry, response_label):
    job_number = entry.get()
    tenant_search = r'G:/' + job_number.replace(job_number[-3:], '000' + '*')
    files = glob.glob(tenant_search, recursive=False)
    if len(files) == 1:
        building_folder = files[0]
        job_folder = glob.glob(building_folder + '/' + job_number[-3:] + '*')
        try:
            os.startfile(job_folder[0])
            response_label.config(text='Job folder opened.')
        except:
            response_label.config(text='Job folder could not be opened with that job number.')
    else:
        print('Could not open folder because more that one result came up.')
        response_label.config(text='Job folder could not be opened with that job number.')


# creates UI using tkinter
def tkinter_ui():
    window = tk.Tk()
    window.title('Project Setup')

    # declarations
    label = tk.Label(text='Insert Job Number', height=2, width=50,)
    entry = tk.Entry(width=40)
    response_label = tk.Label(text='', fg='red', height=2)
    previous_jobs_button = tk.Button(text='Compile Previous Jobs', height=2, relief='raised',
                                     command=lambda:previous_designs(entry, response_label, window))
    open_button = tk.Button(text='Open Project', height=2, relief='raised',
                            command=lambda:open_job_folder(entry, response_label))

    # positions
    label.grid(row=0, columnspan=2)
    entry.grid(row=1, columnspan=2)
    previous_jobs_button.grid(row=2, column=0, pady=20)
    open_button.grid(row=2, column=1, pady=20)
    response_label.grid(row=3, columnspan=2)

    # event watch
    window.mainloop()


#def change_label(response_label):
#    response_label.config(text='Program is processing...')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    tkinter_ui()