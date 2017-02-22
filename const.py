# 1 = only crash errors
# 2 = error + warning
# 3 = All output
error_level = 3

repository = r"C:\tmp\VBad"

#functions available : onClose, onOpen
auto_function_macro = "onOpen"
trigger_close_test_value="True"
trigger_close_test_name = "toto"

#methods available: variables
key_hiding_method = "variable"

#doc_variable options
add_fake_keys = 1
small_keys = 4
big_keys = 3

#encryption available : xor
encryption_type = "xor"
encryption_key_length = 50000 #Max is 65280 for Document.Variable method

#Regex
variable_name_ex = "toto"
regex_rand_var = '\[rdm::([0-9]+)\](\w*)' #regex that select the name of the variable, after the delimiter and the length
regex_rand_del = '\[rdm::[0-9]+\]' #regex should select only the delimiter

regex_defaut_string = '"((?:""|[^"])*)"' #Regex selecting all strings in double quotes (including two consecutives double quotes)
regex_exclude_string_del = '\[!!\]' #The exclusion is to avoid vba string that could finish with exclude characters.
exclude_mark = '[!!]'

regex_string_to_hide = '\[var::(\w*)\]'
regex_string_to_hide_find = '\[var::'+variable_name_ex+'\]'

#Office informations
template_file = repository+r"\Example\Template\template.xls" #Path to the template file used for generate malicious files (To be modified)
filename_list = repository+r"\Example\Lists\filename_list.txt" #Path to the list that contains the filename of the malicious files that will be generated (To be modified)

#saving informations
path_gen_files = repository+r"\Example\Results" #Path were results will be saved (To be modified)

#Malicious VBS Information:
#All data you want to encrypt and include in your doc
original_vba_file = repository+r"\Example\Orignal_VBA\original_vba_prepared.vbs" #Path the prepared VBA files (To be modified)
trigger_function_name =  "Test" #Function that you want to auto_trigger (in your original_vba_file) (To be modified)
string_to_hide = {"domain_name":"http://www.test.com", "path_to_save":r"C:\tmp\toto"}
