import pandas as pd
from labels import lision_label
from labels import lision_custom_pn_label
from labels import lision_outterbox_label
from labels import psi_label
from labels import UMEC_label
from labels import UMEC_outterbox_label
from labels import BL_label
from labels import ADDA_label
from labels import B79_label
from labels import AMERTEK_outterbox_label
from labels import AMERTEK_label
from labels import TZENG_DAR_label
from labels import B75_label
from labels import B75_outterbox_label


class LabelPrinter:
    def __init__(self, excel_file_path: str):
        self.excel_file_path = excel_file_path
        self.data = pd.read_excel(excel_file_path, sheet_name=None, header=0)
        self.param_table = self.data['param_table'].set_index('label_type')
        self.label_status = {label_type: {'total': len(df), 'printed': 0} for label_type, df in self.data.items() if
                             label_type != 'param_table'}

    def print_progress(self, label_type):
        total_labels = self.label_status[label_type]['total']
        total_printed = self.label_status[label_type]['printed']
        print(f"{label_type} progress: {total_printed} out of {total_labels} labels have been printed.")

    def print_test_label(self, label_type: str, label_data: dict):
        # Create a copy of label_data and set the label_copies to 1 for testing
        test_label_data = label_data.copy()
        test_label_data['label_copies'] = 1

        # Print the test label
        self.print_label(label_type, test_label_data, is_test=True)

    def print_label(self, label_type: str, parameters: dict, is_test=False):
        # Define a dictionary that maps label types to their corresponding print functions
        label_print_functions = dict(lision_label=lision_label.modify_and_print_lision_label,
                                     lision_outterbox_label=lision_outterbox_label.modify_and_print_lision_outterbox_label,
                                     lision_custom_pn_label=lision_custom_pn_label.modify_and_print_lision_custom_pn_label,
                                     psi_label=psi_label.modify_and_print_psi_label,
                                     UMEC_label=UMEC_label.modify_and_print_UMEC_label,
                                     UMEC_outterbox_label=UMEC_outterbox_label.modify_and_print_UMEC_outterbox_label,
                                     BL_label=BL_label.modify_and_print_BL_label,
                                     ADDA_label=ADDA_label.modify_and_print_ADDA_label,
                                     B79_label=B79_label.modify_and_print_B79_label,
                                     TZENG_DAR_label = TZENG_DAR_label.modify_and_print_TZENG_DAR_label,
                                     AMERTEK_outterbox_label=AMERTEK_outterbox_label.modify_and_print_AMERTEK_outterbox_label,
                                     AMERTEK_label=AMERTEK_label.modify_and_print_AMERTEK_label,
                                     B75_label=B75_label.modify_and_print_B75_label,
                                     B75_outterbox_label=B75_outterbox_label.modify_and_print_B75_outterbox_label)

        print(f"Printing {label_type} label with parameters {parameters}...")

        # Get the label address for the given label type
        label_address = self.param_table.loc[label_type, 'label_address']

        # Get the print function for the given label
        print_function = label_print_functions.get(label_type)

        if print_function is not None:
            try:
                # Always execute the print operation, regardless of whether this is a test print
                print_function(parameters, label_address)

                # Only update self.label_status if this is not a test print
                if not is_test:
                    self.label_status[label_type]['printed'] += 1
            except Exception as e:
                print(f"Error occurred while printing label: {str(e)}")

    def get_user_input(self,prompt: str) -> str:
        while True:
            choice = input(prompt)
            if choice in ['1', '2']:
                return choice
            else:
                print("Invalid input. Please enter 1 or 2.")

    def perform_label_testing(self, valid_label_types):
        for label_type in valid_label_types:
            # Get the first row of data for this label type
            first_row = self.data[label_type].iloc[0]
            label_data = first_row.to_dict()

            # Print a test label and ask the user if they want to continue
            self.print_test_label(label_type, label_data)

    def print_actual_labels(self, valid_label_types):
        for label_type in valid_label_types:
            df = self.data[label_type]

            # Group the data by the first column
            grouped_data = df.groupby(df.columns[0], sort=False)

            # Iterate over each group
            for group_name, group_df in grouped_data:
                # Iterate over each row of data for this group
                for _, row in group_df.iterrows():
                    # Convert the row to a dictionary
                    label_data = row.to_dict()

                    # Print the actual label
                    self.print_label(label_type, label_data)
                    # After each group is processed, print the progress

            self.print_progress(label_type)

    def preprocess_data(self):
        valid_label_types = []
        for label_type, df in self.data.items():
            if label_type != 'param_table':
                label_address = self.param_table.loc[label_type, 'label_address']
                if not pd.isna(label_address) and not df.empty:
                    valid_label_types.append(label_type)
        return valid_label_types

    def run(self):
        # Preprocess to determine valid label types
        valid_label_types = self.preprocess_data()

        # Ask the user if they want to perform label testing
        choice = self.get_user_input("Do you want to perform label testing? Enter 1 to perform testing, 2 to skip: ")
        if choice == '1':
            self.perform_label_testing(valid_label_types)

        # Ask the user if they want to print actual labels
        choice = self.get_user_input("Do you want to print actual labels? Enter 1 to print, 2 to skip: ")
        if choice == '1':
            self.print_actual_labels(valid_label_types)


if __name__ == "__main__":
    printer = LabelPrinter(r'C:\Users\windows user\Desktop\label_print_20230724.xlsm')
    printer.run()
