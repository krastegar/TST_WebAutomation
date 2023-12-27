import os
import logging


class Menu:
    def __init__(self) -> None:
        pass

    def get_input(self, prompt, expected_type, validate_func=None):
        while True:
            user_input = input(prompt).strip()
            if expected_type == str and prompt.startswith("Enter TST Microsoft Access file_name"):
                if not user_input.endswith(".accdb"):
                    print("Error: file_name must end with .accdb")
                    continue
            try:
                user_input = expected_type(user_input)
            except ValueError as ve:
                logging.exception(f"Error: expected {expected_type.__name__}", ve)
                print(f"Error: expected {expected_type.__name__}")
                continue
            if validate_func and not validate_func(user_input):
                logging.info(f'Invalid Path: {user_input}')
                print(f'Invalid Path: {user_input}')
                continue
            return user_input

    def is_valid_folder_path(self, path):
        """
        Checks if the given string represents a valid folder path and
        contains a file with a .accdb extension.
        Returns True if the path is valid and contains the file, False otherwise.
        """
        if not isinstance(path, str):
            return False
        
        # Use os.path.isdir() to check if the path exists and is a directory
        if not os.path.isdir(path):
            return False
        
        # Use os.listdir() to get a list of files in the directory
        files = os.listdir(path)
        
        # Check if there is at least one file with a .accdb extension
        accdb_files = [f for f in files if f.endswith('.xlsx')]
        
        if not accdb_files:
            print("No Message Monitor exports found...")
            return False
        
        # The folder path is valid and contains at least one .accdb file
        return True


    def get_test_centers(self):
        
        test_centers : list = []
        center_num : int = 0
        yes_list : tuple = ('y', 'yes', 'Y', 'Yes')
        no_list : tuple = ('n', 'no', 'No', 'N')
        accepted_answers : tuple = yes_list+no_list

        # Main code for test centers search only is capable of doing 5 at a time 
        # This can easily be changed in Completeness.py
        while center_num < 5:
            center_num += 1
            test_center : str = self.get_input(f"Enter test_center_{center_num}: ", str)
            test_centers.append(test_center)

            # telling user that we cant have more than 5 inputs
            if center_num == 5:
                print('No more test centers can be added')
                break

            # ask users if they would like to add another test center they are searching for 
            # if they say yes, we repeat the operation, if they say no we exit
            # Exceptions are dealt with in the else portion
            add_center : str = input("Would you like to add another testing center? (y/n): ")
            if add_center in yes_list:
                continue
            elif add_center in no_list: 
                break
            else: 
                # if there is an input that is not yes / no code will repeat until we receive accpepted answer
                while add_center not in accepted_answers:
                    print(f'{add_center} is not a valid input')
                    add_center = input("Would you like to add another testing center? (y/n): ")

                    # if answer is no, then we break out of both while loops
                    # if the answer is yes, we only break out of inner while loop 
                    if add_center in no_list:
                        break
                if add_center in no_list:
                    break
                continue

        # Cannot have more than 5 centers currently         
        assert len(test_centers) <= 5

        return test_centers