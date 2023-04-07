import mdPresort_pythoncode
import os
import sys


class DataContainer:
    def __init__(self, presort_file="", output_file=""):
        self.presort_file = presort_file
        self.output_file = output_file

    def format_presort_output_file(self):
        # Change the name of the output file
        location = self.presort_file.find(".csv")
        self.output_file = self.presort_file[:location] + "_output.csv"

    def get_wrapped(self):
        file = os.path.abspath(self.output_file)
        file_path = file.replace("\\", "/")

        max_line_length = 50
        lines = file_path.split("/")
        current_line = ""
        wrapped_string = []

        for section in lines:
            if len(current_line + section) > max_line_length:
                wrapped_string.append(current_line.strip())
                current_line = ""

            if self.output_file in section:
                current_line += section
            else:
                current_line += section + "/"

        if len(current_line) > 0:
            wrapped_string.append(current_line.strip())

        return wrapped_string



class PresortObject:
    """ Set license string and set path to data files  (.dat, etc) """

    def __init__(self, license, data_path):
        self.md_presort_obj = mdPresort_pythoncode.mdPresort()
        self.md_presort_obj.SetLicenseString(license)
        self.data_path = data_path
        self.md_presort_obj.SetPathToPresortDataFiles(data_path)

        """
        If you see a different date than expected, check your license string and either download the new data files or use the Melissa Updater program to update your data files.  
           
        If 1970-00-00 is the DatabaseDate, the Phone Object was unable to reach the data files.
        """

        p_status = self.md_presort_obj.InitializeDataFiles()
        if (p_status != mdPresort_pythoncode.ProgramStatus.ErrorNone):
            print("Failed to Initialize Object.")
            print(p_status)
            return

        print(
            f"                DataBase Date: {self.md_presort_obj.GetDatabaseDate()}")
        print(
            f"              Expiration Date: {self.md_presort_obj.GetLicenseStringExpirationDate()}")

        """
        This number should match with file properties of the Melissa Object binary file.
        If TEST appears with the build number, there may be a license key issue.
        """
        print(
            f"               Object Version: {self.md_presort_obj.GetBuildNumber()}\n")

    def execute_object_and_result_codes(self, data):
        """
        PreSortSettings to the desired type of Presort
        The sample defaults to First Class Mailing, if you'd like to use Standard, comment out the current line and uncomment the one containing the enumeration: STD_LTR_AUTO
        STD_LTR_AUTO is for Standard Mail, Letter, Automation pieces
        FCM_LTR_AUTO = First Class Mail, Letter, Automation pieces
        """
        self.md_presort_obj.SetPreSortSettings(
            int(mdPresort_pythoncode.SortationCode.FCM_LTR_AUTO.value))

        # self.md_presort_obj.SetPreSortSettings(2)

        # self.md_presort_obj.PreSortSettings(PresortObjLib.SortationCode.STD_LTR_AUTO)

        # Sack and Parcel Dimensions
        self.md_presort_obj.SetSackWeight("30")
        self.md_presort_obj.SetPieceLength("9")
        self.md_presort_obj.SetPieceHeight("4.5")
        self.md_presort_obj.SetPieceThickness("0.042")
        self.md_presort_obj.SetPieceWeight("1.5")

        """
        Mailers ID
        Insert a valid 6-9 digit MailersID number
        If you do not have a valid Mailers ID, you can visit the USPS to apply for one. Go to usps.com and click on the 'Business Customer Gateway' link at the bottom of the page.
        """
        self.md_presort_obj.SetMailersID("123456")

        """
        Postage Payment Methods
        If both of these functions are set to false (Default setting), then Metered Mail will be used.
        Permit Imprint - Mailing without affixing postage. When set to true, current mailing will use permit imprint as method of postage payment
        """
        # self.md_presort_obj.PSPermitImprint = false; # way the mailer pays for his mailing

        """
        Precanceled Stamp - Cancellation of adhesive postage, stamped envelopes, or stamped cards before mailing.
        When set to true, PSPrecanceledStampValue must be set to the postage value in cents
        """
        # self.md_presort_obj.PSPrecanceledStamp = true;
        # self.md_presort_obj.PSPrecanceledStampValue(0.10);

        """
        Sorting Options
        PSNCOA - Uses NCOALink for move updates
        """

        self.md_presort_obj.SetPSNCOA(True)

        # Post Office of Mailing Information - This is the Post office from where it will be mailed from
        self.md_presort_obj.SetPSPostOfficeOfMailingCity("RSM")
        self.md_presort_obj.SetPSPostOfficeOfMailingState("CA")
        self.md_presort_obj.SetPSPostOfficeOfMailingZIP("92688")

        """
        Update the Parameters
        Verify and validate that the following properties were set to correct ranges: SetPieceHeight, SetPieceLength, SetPieceThickness,and SetPieceWeight
        """
        if not self.md_presort_obj.UpdateParameters():
            print("Parameter Error: " +
                  self.md_presort_obj.GetParametersErrorString())
            return
        """
        Add Input Records to Presort Object
        Parse through the input file and add each record to the Presort Object
        """
        address_dict = {}
        debugcount = 0
        try:
            with open(data.presort_file) as reader:
                header_row = reader.readline()  # Header row
                line = reader.readline()
                while line:
                    split = line.split(',')

                    self.md_presort_obj.SetRecordID(split[0])
                    
                    self.md_presort_obj.SetZip(split[6])
                    self.md_presort_obj.SetPlus4(split[7])
                    self.md_presort_obj.SetDeliveryPointCode(split[8])
                    self.md_presort_obj.SetCarrierRoute(split[9].strip())

                    address_info = ','.join(split[1:10]).strip()
                    address_dict[split[0]] = address_info


                    if self.md_presort_obj.AddRecord() == False:
                        print(
                            f"\nError Adding Record {split[0]}: {self.md_presort_obj.GetParametersErrorString()}")

                    line = reader.readline()

        except Exception as ex:
            print(f"\nError Reading File: {str(ex)}")

        finally:
            if reader != None:
                reader.close()

        # Do Presort - Do and check if Presort was successful, otherwise return the error.
        if self.md_presort_obj.DoPresort() == False:
            print(
                f"Error During Presort: {self.md_presort_obj.GetParametersErrorString()}")

        # Write Results
        try:
            with open(data.output_file, 'w') as f:
                # Write the headers on the first line
                f.write(
                    "RecID,MAK,Address,Suite,City,State,Zip,Plus4,DeliveryPointCode,CarrierRoute,TrayNumber,SequenceNumber,Endorsement,BundleNumber\n")

                # Grab the first record and continue through the file
                writeback = self.md_presort_obj.GetFirstRecord()

                while writeback:
                    f.write(f"{self.md_presort_obj.GetRecordID()}," +
                            f"{address_dict[self.md_presort_obj.GetRecordID()]}," +
                            f"{self.md_presort_obj.GetTrayNumber()},{self.md_presort_obj.GetSequenceNumber()}," + 
                            f"{self.md_presort_obj.GetEndorsementLine()},{self.md_presort_obj.GetBundleNumber()}" +
                            "\n")
                    writeback = self.md_presort_obj.GetNextRecord()
        except Exception as ex:
            print("\nError Writing File:", ex)

        return DataContainer(data.presort_file, data.output_file)


def parse_arguments():
    license, test_presort_file, data_path = "", "", ""

    args = sys.argv
    index = 0
    for arg in args:

        if (arg == "--license") or (arg == "-l"):
            if (args[index+1] != None):
                license = args[index+1]
        if (arg == "--dataPath") or (arg == "-d"):
            if (args[index+1] != None):
                data_path = args[index+1]
        if (arg == "--file") or (arg == "-f"):
            if (args[index+1] != None):
                test_presort_file = args[index+1]
        index += 1

    return (license, test_presort_file, data_path)


def run_as_console(license, test_presort_file, data_path):
    print("\n\n======== WELCOME TO MELISSA PRESORT OBJECT WINDOWS PYTHON3 =========\n")

    presort_object = PresortObject(license, data_path)

    should_continue_running = True

    if presort_object.md_presort_obj.GetInitializeErrorString() != "No Errors":
        should_continue_running = False

    while should_continue_running:
        if test_presort_file == None or test_presort_file == "":
            print("\nFill in each value to see the Presort Object results")
            presort_file = str(input("Presort file path: "))
        else:
            presort_file = test_presort_file

        data = DataContainer(presort_file)
        data.format_presort_output_file()

        """ Print user input """
        print("\n============================== INPUTS ==============================\n")
        print(f"\t         Presort File: {presort_file}")

        """ Execute Phone Object """
        data_container = presort_object.execute_object_and_result_codes(data)

        """ Print output """
        print("\n============================== OUTPUT ==============================\n")
        print("\n\tPresort Object Information:")
        sections = data_container.get_wrapped()

        print(f"\t          Output File: {sections[0]}")

        for i in range(1, len(sections)):
            if i == len(sections) - 1 and sections[i].endswith("/"):
                sections[i] = sections[i][:-1]
            print(f"\t                       {sections[i]}")


        is_valid = False
        if not (test_presort_file == None or test_presort_file == ""):
            is_valid = True
            should_continue_running = False
        while not is_valid:

            test_another_response = input(str("\nTest another file? (Y/N)\n"))

            if not (test_another_response == None or test_another_response == ""):
                test_another_response = test_another_response.lower()
            if test_another_response == "y":
                is_valid = True

            elif test_another_response == "n":
                is_valid = True
                should_continue_running = False
            else:

                print("Invalid Response, please respond 'Y' or 'N'")

    print("\n============= THANK YOU FOR USING MELISSA PYTHON3 OBJECT ===========\n")


"""  MAIN STARTS HERE   """

license, test_presort_file, data_path = parse_arguments()

run_as_console(license, test_presort_file, data_path)
