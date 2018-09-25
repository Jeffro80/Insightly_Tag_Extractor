# -*- coding: utf-8 -*-
# Insightly Status Tag Extractor
# Version 1.0 25 September 2018
# Created by Jeff Mitchell
# Pulls out the status tag for each student from Insightly Data dump
# Insightly Report is 'Contact Tag List'

# To Do:

# To Fix:



import custtools.admintools as ad
import custtools.filetools as ft
import re
import sys
import pandas as pd


def check_existing_students(report_data):
    """Return list of warnings for information in Existing students data.

    Checks the Existing students data to see if the required information is
    present. Missing or incorrect information that is non-fatal is appended to
    a warnings list and returned.

    Args:
        report_data (list): Existing students report data.

    Returns:
        True if warnings list has had items appended to it, False otherwise.
        warnings (list): Warnings that have been identified in the data.

    File structure (report_data):
        EnrolmentPK, StudentID, CourseFK, TutorFK, StartDate, ExpiryDate,
        Status, Tag.

    File source (report_data):
        Enrolments Table in Student Database.
    """
    errors = []
    i = 0
    warnings = ['\nExisting Students Data File Warnings:\n']
    while i < len(report_data):
        student = report_data[i]
        if student[1] in (None, ''):
            errors.append('Student ID is missing for student with the '
                            'Enrolment Code {}'.format(student[0]))
        if student[6] in (None, ''):
            warnings.append('Status is missing for student with the '
                            'Student ID {}'.format(student[1]))
        if student[7] in (None, ''):
            warnings.append('Tag is missing for student with the '
                            'Student ID {}'.format(student[1]))
        # Add check that tag is correct
        if student[7].lower() not in ('n/a', 'green', 'orange', 'red', 'black',
                  'purple', 'suspended', 'withdrawn', 'graduated', 'expired',
                  'on hold', 'cancelled'):
            errors.append('Tag for student with the Student ID {} is not '
                          'valid.'.format(student[1]))
        i += 1
    # Check if any errors have been identified, save error log if they have
    if len(errors) > 0:
        ft.process_error_log(errors, 'Existing Students Data File')
    # Check if any warnings have been identified, save error log if they have
    if len(warnings) > 1:
        return True, warnings
    else:
        return False, warnings


def check_insightly(report_data):
    """Return list of warnings for information in Insightly data.

    Checks the Insightly data to see if the required information is present.
    Missing or incorrect information that is non-fatal is appended to a
    warnings list and returned.

    Args:
        report_data (list): Insightly report data.

    Returns:
        True if warnings list has had items appended to it, False otherwise.
        warnings (list): Warnings that have been identified in the data.

    File structure (report_data):
        StudentID, First Name, Last Name, Tags.

    File source (report_data):
        Insightly Data Dump (using columns listed in File structure).
    """
    errors = []
    i = 0
    warnings = ['\nInsightly Data File Warnings:\n']
    while i < len(report_data):
        student = report_data[i]
        if student[1] in (None, ''):
            errors.append('First Name is missing for student with the '
                          'Student ID {}'.format(student[0]))
        if student[2] in (None, ''):
            errors.append('Last Name is missing for student with the '
                          'Student ID {}'.format(student[0]))
        if student[3] in (None, ''):
            warnings.append('Tags is missing for student with the '
                          'Student ID {}'.format(student[0]))
        i += 1
    # Check if any errors have been identified, save error log if they have
    if len(errors) > 0:
        ft.process_error_log(errors, 'Insightly Data File')
    # Check if any warnings have been identified, save error log if they have
    if len(warnings) > 1:
        return True, warnings
    else:
        return False, warnings


def check_repeat():
    """Return True or False for repeating another action.

    Returns:
        True if user wants to perform another action, False otherwise.
    """
    repeat = ''
    while repeat == '':
        repeat = input("\nDo you want to prepare another file? y/n --> ")
        if repeat != 'y' and repeat != 'n':
            print("\nThat is not a valid answer! Please try again.")
            repeat = ''
        elif repeat == 'y':
            return True
        else:
            return False


def check_review_warnings():
    """Return True or False for reviewing warning messages.

    Returns:
        True if user wants to review warning messages, False otherwise.
    """
    review = ''
    while review == '':
        review = input('\nDo you want to view the warning messages? y/n --> ')
        if review not in ('y', 'n'):
            print('\nThat is not a valid answer! Please try again.')
            review = ''
        elif review == 'y':
            return True
        else:
            return False


def clean_exist_stud(raw_data):
    """Clean data in the Existing students data.
    
    Extracts the desired columns from the Existing student data and cleans the
    raw data.
    
    Args:
        raw_data (list): Raw existing student data.
        
    Returns:
        cleaned_data (list): Existing student data that has been cleaned.
    
    File structure (raw_data):
        EnrolmentPK, StudentID, CourseFK, TutorFK, StartDate, ExpiryDate,
        Status, Tag.

    File source (raw_data):
        Enrolments Table in Student Database.
    """
    cleaned_data = []
    i = 0
    while i < len(raw_data):
        student = raw_data[i]
        cleaned_student = []
        # Extract and clean desired columns
        cleaned_student.append(student[0].strip())
        cleaned_student.append(student[1].strip())
        cleaned_student.append(student[7].strip())
        cleaned_data.append(cleaned_student)
        i += 1
    return cleaned_data        


def clean_insightly(raw_data):
    """Clean data in the Insightly data.
    
    Extracts the desired columns from the Insightly data and cleans the raw
    data.
    
    Args:
        raw_data (list): Raw insightly data.
        
    Returns:
        cleaned_data (list): Insightly data that has been cleaned.
    
    File structure (report_data):
        StudentID, First Name, Last Name, Tags.

    File source (report_data):
        Insightly Data Dump (using columns listed in File structure).
    """
    cleaned_data = []
    i = 0
    while i < len(raw_data):
        student = raw_data[i]
        cleaned_student = []
        # Extract and clean desired columns
        cleaned_student.append(student[0].strip())
        cleaned_student.append(student[1].strip())
        cleaned_student.append(student[2].strip())
        cleaned_student.append(student[3].strip())
        cleaned_data.append(cleaned_student)
        i += 1
    return cleaned_data


def extract_course_tag(raw_data, courses):
    """Replace Contact tag with Course tag.
    
    Replaces the Tags field with the student's Course tag, extracted
    from the Tags list.
    
    Args:
        raw_data (str): String containing Contact tag data.
        courses (list): List of course codes to check for.
        
    Returns:
        The course code tag if found, 'N/A' otherwise.
    """
    for course in courses:
        if re.search(course.lower(), raw_data.lower()):
            return course
    # No valid course found
    return 'N/A'


def extract_status_tag(raw_data):
    """Replace Contact tag with Status tag.
    
    Replaces the Tags field with the colour tag for their status, extracted
    from the Tags list.
    
    Args:
        raw_data (str): String containing Contact tag data.
        
    Returns:
        The colour of their status tag if found, 'N/A' otherwise.
    """
    if re.search('suspended', raw_data.lower()):
        return 'Suspended'
    elif re.search('withdrawn', raw_data.lower()):
        return 'Withdrawn'
    elif re.search('graduated', raw_data.lower()):
        return 'Graduated'
    elif re.search('expired', raw_data.lower()):
        return 'Expired'
    elif re.search('on hold', raw_data.lower()):
        return 'On Hold'
    elif re.search('cancelled', raw_data.lower()):
        return 'Cancelled'
    elif re.search('green', raw_data.lower()):
        return 'Green'
    elif re.search('orange', raw_data.lower()):
        return 'Orange'
    elif re.search('red', raw_data.lower()):
        return 'Red'
    elif re.search('black', raw_data.lower()):
        return 'Black'
    elif re.search('purple', raw_data.lower()):
        return 'Purple'
    else:
        return 'N/A'


def extract_tutor_tag(raw_data, tutors):
    """Replace Contact tag with Tutor tag.
    
    Replaces the Tags field with the student's Tutor tag, extracted
    from the Tags list.
    
    Args:
        raw_data (str): String containing Contact tag data.
        tutors (list): List of tutor names to check for.
        
    Returns:
        The name of their tutor tag if found, 'N/A' otherwise.
    """
    for tutor in tutors:
        if re.search(tutor.lower(), raw_data.lower()):
            return tutor
    # No valid tutor found
    return 'N/A'


def get_old_response():
    """Return user input for inclusion of old students.
    
    Returns:
        old (bool): User selection for including old students.
    """
    response = ''
    while response == '':
        response = input('\nDo you want to include all students, including '
                         'those that are not in the Student Database? y/n '
                         '--> ')
        if response.lower() not in ['y', 'n']:
            print('\nThat is not a valid answer! Please try again.')
            response = ''
        elif response.lower() == 'y':
            return True
        else:
            return False


def get_sample():
    """Return user input for source of data.
    
    Returns:
        source (str): User selection for source of data.
    """
    repeat = True
    high = 2
    while repeat:
        sample_menu()
        try:
            action = int(input('\nPlease enter the number for your '
                               'selection --> '))
        except ValueError:
            print('Please enter a number between 1 and {}.'.format(high))
        else:
            if action < 1 or action > high:
                print('\nPlease select from the available options (1 - {})'
                      .format(high))
            elif action == 1:
                return 'All'
            elif action == 2:
                return 'Active'


def load_data(source, f_name=''):
    """Read data from a file.

    Args:
        source (str): The code for the table that the source data belongs to.
        f_name (str): (Optional) File name to be loaded. If not provided, user
        will be prompted to provide a file name.

    Returns:
        read_data (list): A list containing the data read from the file.
        True if warnings list has had items appended to it, False otherwise.
        warnings (list): Warnings that have been identified in the data.
    """
    warnings = []
    # Load file
    if f_name in (None, ''): # Get from user
        read_data = ft.get_csv_fname_load(source)
    else:
        read_data = ft.load_csv(f_name, 'e')
    # Check that data has entries for each required column
    if source in ['All Students Data', 'Active Students Data']:
        to_add, items_to_add = check_existing_students(read_data)
        if to_add:
            for item in items_to_add:
                warnings.append(item)
    elif source == 'Insightly_Data_':
        to_add, items_to_add = check_insightly(read_data)
        if to_add:
            for item in items_to_add:
                warnings.append(item)
    if len(warnings) > 0:
        return read_data, True, warnings
    else:
        return read_data, False, warnings


def main():
    repeat = True
    high = 5
    while repeat is True:
        try_again = False
        main_message()
        try:
            action = int(input('\nPlease enter the number for your '
                               'selection --> '))
        except ValueError:
            print('Please enter a number between 1 and {}.'.format(high))
            try_again = True
        else:
            if action < 1 or action > high:
                print('\nPlease select from the available options '
                      '(1 - {}).'.format(high))
                try_again = True
            elif action == 1:
                process_status_tag_extraction()
            elif action == 2:
                process_tutor_tag_extraction()
            elif action == 3:
                process_course_tag_extraction()
            elif action == 4:
                process_all_tags_extraction()
            elif action == 5:
                print('\nIf you have generated any files, please find them '
                      'saved to disk. Goodbye.')
                sys.exit()
        if not try_again:
            repeat = check_repeat()
    print('\nPlease find your files saved to disk. Goodbye.')


def main_message():
    """Print the menu of options."""
    print('\n\n*************==========================*****************')
    print('\nInsightly_Tag_Extractor version 1.0')
    print('Created by Jeff Mitchell, 2018')
    print('\nOptions:')
    print('\n1. Extract Status Tags')
    print('2. Extract Tutor Tags')
    print('3. Extract Course Tags')
    print('4. Extract All Tags')
    print('5. Exit')


def old_menu():
    """Display the old students check menu options."""
    print('\nPlease enter the number for the source of the data:\n')
    print('1: All Students')
    print('2: Active Students')


def process_all_tags_extraction():
    """Process all tags for extraction.
    
    Extracts the course, tutor and status tag for each student. Returns a
    DataFrame with the extracted students.
    
    File structure (Existing students):
        EnrolmentPK, StudentID, CourseFK, TutorFK, StartDate, ExpiryDate,
        Status, Tag.
        
    File structure (Insightly_Data):
        StudentID, First Name, Last Name, Tags.
        
    File structure (courses.txt):
       Code of each course separated by a comma (no spaces).
    
    File structure (tutors.txt):
        First name of each tutor separated by a comma (no spaces).
        
    File source (Existing students):
        Enrolments Table in Student Database.
        
    File source (Insightly_Data):
        Insightly Data Dump (using columns listed in File structure).
    
    File source (courses.txt):
        Course codes taken from Student Database.
    
    File source (tutors.txt):
        Tutors in Insightly (check Contact Tags in Contacts).
    """
    warnings = ['\nProcessing All Student Tags Extraction data Warnings:\n']
    warnings_to_process = False
    print('\nExtracting All Student Tags.')
    # Confirm the required files are in place
    required_files = ['Existing Students', 'Insightly Data', 'Course Codes',
                      'Tutor Tags']
    ad.confirm_files('Extracting All Student Tags', required_files)
    # Ask if want all students or only those in the Student Database
    keep_old = get_old_response()
    # Get sample to use - if do not want old students
    if keep_old:
        source = 'All Students Data'
        sample = 'All'
    else:
        sample = get_sample()
        source = '{} Students Data'.format(sample)
    print('\nYou will need to load the {} file.'.format(source))
    # Get name for 'Existing Students' Report data file and then load
    exist_stud_data, to_add, warnings_to_add = load_data(source)
    if to_add:
        warnings_to_process = True
        for line in warnings_to_add:
            warnings.append(line)
    # Get name for 'Insightly_Data' Report data file and then load
    insightly_data, to_add, warnings_to_add = load_data('Insightly_Data_')
    if to_add:
        warnings_to_process = True
        for line in warnings_to_add:
            warnings.append(line)
    # Load Courses File
    courses = ft.load_headings('courses.txt')
    # Load Tutors File
    tutors = ft.load_headings('tutors.txt')
    # Clean the Existing Student data and extract desired columns
    exist_stud_clean = clean_exist_stud(exist_stud_data)
    # Clean the Insightly data and extract desired columns
    insightly_clean = clean_insightly(insightly_data)
    # Create DataFrame for Existing students
    headings = ['Enrolment Code', 'StudentID', 'Tag']
    exist = pd.DataFrame(data = exist_stud_clean, columns = headings)
    # Create DataFrame for Insightly Data
    headings = ['StudentID', 'First Name', 'Last Name', 'Tags']
    insight = pd.DataFrame(data = insightly_clean, columns = headings)
    # Remove Expired, Graduated and Withdrawn students if desired
    if sample == 'Active':
        insight['Tags'] = insight['Tags'].apply(remove_inactive)
        insight = insight.drop(insight.index[insight['Tags'] == 
                                             'Remove'])
    # Find course tag and save to column
    insight['Course'] = insight['Tags'].apply(extract_course_tag,
           args=(courses,))
    # Find status tag and save to column
    insight['Status'] = insight['Tags'].apply(extract_status_tag)
    # Find tutor tag and save to column
    insight['Tutor'] = insight['Tags'].apply(extract_tutor_tag,
           args=(tutors,))
    # Remove Students not in the Student Database, if required
    if keep_old: # Keep all students in Insightly records
        tags = pd.merge(exist, insight, on='StudentID', how='right')
        tags['Enrolment Code'] = tags['Enrolment Code'].fillna('N/A')
    else: # Remove students not in the Student Database
        tags = pd.merge(exist, insight, on='StudentID', how='inner')
    headings = ['Enrolment Code', 'StudentID', 'First Name', 'Last Name',
                'Course', 'Tutor', 'Status']
    tags = tags[headings]
    # Save Master file
    f_name = 'All_Tags_{}.xls'.format(ft.generate_time_string())
    tags.to_excel(f_name, index=False)
    print('\nAll_Tags has been saved to {}'.format(f_name))
    ft.process_warning_log(warnings, warnings_to_process)


def process_course_tag_extraction():
    """Process course tag extraction.
    
    Extracts the course tag for each student. Returns a DataFrame with the
    extracted students.
    
    File structure (Existing students):
        EnrolmentPK, StudentID, CourseFK, TutorFK, StartDate, ExpiryDate,
        Status, Tag.
        
    File structure (Insightly_Data):
        StudentID, First Name, Last Name, Tags.
        
    File structure (courses.txt):
       Code of each course separated by a comma (no spaces).
        
    File source (Existing students):
        Enrolments Table in Student Database.
        
    File source (Insightly_Data):
        Insightly Data Dump (using columns listed in File structure).
    
    File source (courses.txt):
        Course codes taken from Student Database.
    """
    warnings = ['\nProcessing Student Course Tag Extraction data Warnings:\n']
    warnings_to_process = False
    print('\nExtracting Student Course Tags.')
    # Confirm the required files are in place
    required_files = ['Existing Students', 'Insightly Data', 'Course Codes']
    ad.confirm_files('Extracting Student Course Tags', required_files)
    # Ask if want all students or only those in the Student Database
    keep_old = get_old_response()
    # Get sample to use - if do not want old students
    if keep_old:
        source = 'All Students Data'
        sample = 'All'
    else:
        sample = get_sample()
        source = '{} Students Data'.format(sample)
    print('\nYou will need to load the {} file.'.format(source))
    # Get name for 'Existing Students' Report data file and then load
    exist_stud_data, to_add, warnings_to_add = load_data(source)
    if to_add:
        warnings_to_process = True
        for line in warnings_to_add:
            warnings.append(line)
    # Get name for 'Insightly_Data' Report data file and then load
    insightly_data, to_add, warnings_to_add = load_data('Insightly_Data_')
    if to_add:
        warnings_to_process = True
        for line in warnings_to_add:
            warnings.append(line)
    # Load Courses File
    courses = ft.load_headings('courses.txt')
    # Clean the Existing Student data and extract desired columns
    exist_stud_clean = clean_exist_stud(exist_stud_data)
    # Clean the Insightly data and extract desired columns
    insightly_clean = clean_insightly(insightly_data)
    # Create DataFrame for Existing students
    headings = ['Enrolment Code', 'StudentID', 'Tag']
    exist = pd.DataFrame(data = exist_stud_clean, columns = headings)
    # Create DataFrame for Insightly Data
    headings = ['StudentID', 'First Name', 'Last Name', 'Tags']
    insight = pd.DataFrame(data = insightly_clean, columns = headings)
    # Remove Expired, Graduated and Withdrawn students if desired
    if sample == 'Active':
        insight['Tags'] = insight['Tags'].apply(remove_inactive)
        insight = insight.drop(insight.index[insight['Tags'] == 
                                             'Remove'])
    # Find course tag and save to column
    insight['Tags'] = insight['Tags'].apply(extract_course_tag,
           args=(courses,))
    # Remove Students not in the Student Database, if required
    if keep_old: # Keep all students in Insightly records
        tags = pd.merge(exist, insight, on='StudentID', how='right')
        tags['Enrolment Code'] = tags['Enrolment Code'].fillna('N/A')
    else: # Remove students not in the Student Database
        tags = pd.merge(exist, insight, on='StudentID', how='inner')
    headings = ['Enrolment Code', 'StudentID', 'First Name', 'Last Name',
                'Tags']
    tags = tags[headings]
    # Save Master file
    f_name = 'Course_Tags_{}.xls'.format(ft.generate_time_string())
    tags.to_excel(f_name, index=False)
    print('\nCourse_Tags has been saved to {}'.format(f_name))
    ft.process_warning_log(warnings, warnings_to_process)


def process_status_tag_extraction():
    """Process status tag extraction.
    
    Extracts the status tag for each student. Returns a DataFrame with the
    extracted students.
    
    File structure (Existing students):
        EnrolmentPK, StudentID, CourseFK, TutorFK, StartDate, ExpiryDate,
        Status, Tag.
        
    File structure (Insightly_Data):
        StudentID, First Name, Last Name, Tags.
        
    File source (Existing students):
        Enrolments Table in Student Database.
        
    File source (Insightly_Data):
        Insightly Data Dump (using columns listed in File structure).
    """
    warnings = ['\nProcessing Student Status Tag Extraction data Warnings:\n']
    warnings_to_process = False
    print('\nExtracting Student Status Tags.')
    # Confirm the required files are in place
    required_files = ['Existing Students', 'Insightly_Data']
    ad.confirm_files('Extracting Student Status Tags', required_files)
    # Ask if want all students or only those in the Student Database
    keep_old = get_old_response()
    # Get sample to use - if do not want old students
    if keep_old:
        source = 'All Students Data'
        sample = 'All'
    else:
        sample = get_sample()
        source = '{} Students Data'.format(sample)
    print('\nYou will need to load the {} file.'.format(source))
    # Get name for 'Existing Students' Report data file and then load
    exist_stud_data, to_add, warnings_to_add = load_data(source)
    if to_add:
        warnings_to_process = True
        for line in warnings_to_add:
            warnings.append(line)
    # debug_list(exist_stud_data)
    # Get name for 'Insightly_Data' Report data file and then load
    insightly_data, to_add, warnings_to_add = load_data('Insightly_Data_')
    if to_add:
        warnings_to_process = True
        for line in warnings_to_add:
            warnings.append(line)
    # debug_list(insightly_data)
    # Clean the Existing Student data and extract desired columns
    exist_stud_clean = clean_exist_stud(exist_stud_data)
    # debug_list(exist_stud_clean)
    # Clean the Insightly data and extract desired columns
    insightly_clean = clean_insightly(insightly_data)
    # debug_list(insightly_clean)
    # Create DataFrame for Existing students
    headings = ['Enrolment Code', 'StudentID', 'Tag']
    exist = pd.DataFrame(data = exist_stud_clean, columns = headings)
    # Create DataFrame for Insightly Data
    headings = ['StudentID', 'First Name', 'Last Name', 'Tags']
    insight = pd.DataFrame(data = insightly_clean, columns = headings)
    # Remove Expired, Graduated and Withdrawn students if desired
    if sample == 'Active':
        insight['Tags'] = insight['Tags'].apply(remove_inactive)
        insight = insight.drop(insight.index[insight['Tags'] == 
                                             'Remove']) 
    # Find status tag and save to column
    insight['Tags'] = insight['Tags'].apply(extract_status_tag)
    # Remove Students not in the Student Database, if required
    if keep_old: # Keep all students in Insightly records
        tags = pd.merge(exist, insight, on='StudentID', how='right')
        tags['Enrolment Code'] = tags['Enrolment Code'].fillna('N/A')
    else: # Remove students not in the Student Database
        tags = pd.merge(exist, insight, on='StudentID', how='inner')
    headings = ['Enrolment Code', 'StudentID', 'First Name', 'Last Name',
                'Tags']
    tags = tags[headings]
    # Save Master file
    f_name = 'Status_Tags_{}.xls'.format(ft.generate_time_string())
    tags.to_excel(f_name, index=False)
    print('\nStatus_Tags has been saved to {}'.format(f_name))
    ft.process_warning_log(warnings, warnings_to_process)


def process_tutor_tag_extraction():
    """Process tutor tag extraction.
    
    Extracts the tutor tag for each student. Returns a DataFrame with the
    extracted students.
    
    File structure (Existing students):
        EnrolmentPK, StudentID, CourseFK, TutorFK, StartDate, ExpiryDate,
        Status, Tag.
        
    File structure (Insightly_Data):
        StudentID, First Name, Last Name, Tags.
        
    File structure (tutors.txt):
        First name of each tutor separated by a comma (no spaces).
        
    File source (Existing students):
        Enrolments Table in Student Database.
        
    File source (Insightly_Data):
        Insightly Data Dump (using columns listed in File structure).
    
    File source (tutors.txt):
        Tutors in Insightly (check Contact Tags in Contacts).
    """
    warnings = ['\nProcessing Student Tutor Tag Extraction data Warnings:\n']
    warnings_to_process = False
    print('\nExtracting Student Tutor Tags.')
    # Confirm the required files are in place
    required_files = ['Existing Students', 'Insightly Data', 'Tutor Names']
    ad.confirm_files('Extracting Student Tutor Tags', required_files)
    # Ask if want all students or only those in the Student Database
    keep_old = get_old_response()
    # Get sample to use - if do not want old students
    if keep_old:
        source = 'All Students Data'
        sample = 'All'
    else:
        sample = get_sample()
        source = '{} Students Data'.format(sample)
    print('\nYou will need to load the {} file.'.format(source))
    # Get name for 'Existing Students' Report data file and then load
    exist_stud_data, to_add, warnings_to_add = load_data(source)
    if to_add:
        warnings_to_process = True
        for line in warnings_to_add:
            warnings.append(line)
    # Get name for 'Insightly_Data' Report data file and then load
    insightly_data, to_add, warnings_to_add = load_data('Insightly_Data_')
    if to_add:
        warnings_to_process = True
        for line in warnings_to_add:
            warnings.append(line)
    # Load Tutors File
    tutors = ft.load_headings('tutors.txt')
    # Clean the Existing Student data and extract desired columns
    exist_stud_clean = clean_exist_stud(exist_stud_data)
    # Clean the Insightly data and extract desired columns
    insightly_clean = clean_insightly(insightly_data)
    # Create DataFrame for Existing students
    headings = ['Enrolment Code', 'StudentID', 'Tag']
    exist = pd.DataFrame(data = exist_stud_clean, columns = headings)
    # Create DataFrame for Insightly Data
    headings = ['StudentID', 'First Name', 'Last Name', 'Tags']
    insight = pd.DataFrame(data = insightly_clean, columns = headings)
    # Remove Expired, Graduated and Withdrawn students if desired
    if sample == 'Active':
        insight['Tags'] = insight['Tags'].apply(remove_inactive)
        insight = insight.drop(insight.index[insight['Tags'] == 
                                             'Remove'])
    # Find tutor tag and save to column
    insight['Tags'] = insight['Tags'].apply(extract_tutor_tag, args=(tutors,))
    # Remove Students not in the Student Database, if required
    if keep_old: # Keep all students in Insightly records
        tags = pd.merge(exist, insight, on='StudentID', how='right')
        tags['Enrolment Code'] = tags['Enrolment Code'].fillna('N/A')
    else: # Remove students not in the Student Database
        tags = pd.merge(exist, insight, on='StudentID', how='inner')
    headings = ['Enrolment Code', 'StudentID', 'First Name', 'Last Name',
                'Tags']
    tags = tags[headings]
    # Save Master file
    f_name = 'Tutor_Tags_{}.xls'.format(ft.generate_time_string())
    tags.to_excel(f_name, index=False)
    print('\nTutor_Tags has been saved to {}'.format(f_name))
    ft.process_warning_log(warnings, warnings_to_process)


def remove_inactive(raw_data):
    """Replace Contact tag for unwanted students.
    
    Replaces the Tags field with 'Remove' if the student contains
    the tag 'Withdrawn', 'Graduated', 'Expired' or 'Transferred'. This allows
    the students to be dropped from the data in a later step.
    
    Args:
        raw_data (str): String containing Contact tag data.
        
    Returns:
        'Remove' if a tag 'Withdrawn', 'Graduated', 'Expired' or 'Transferred'
        is found, the passed Tags otherwise.
    """
    if re.search('withdrawn', raw_data.lower()):
        return 'Remove'
    elif re.search('expired', raw_data.lower()):
        return 'Remove'
    elif re.search('graduated', raw_data.lower()):
        return 'Remove'
    elif re.search('transferred', raw_data.lower()):
        return 'Remove'
    else:
        return raw_data


def sample_menu():
    """Display the sample menu options."""
    print('\nWill you be processing All students in the Student Database '
          'or just those that are currently Active?:\n')
    print('1: All Students')
    print('2: Active Students')


if __name__ == '__main__':
    main()