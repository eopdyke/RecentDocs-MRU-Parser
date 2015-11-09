import _winreg
import sys
import subprocess
from datetime import datetime



def help_cmd():
    unload()
    sys.exit('\n[X] Must be run from an Administrative cmd prompt\n[+] usage: mru_parse.py path_to_NTUSER.dat_file [path to output file]')
    
def load():
    try:
        print '\n[+] Loading '+sys.argv[1]+ ' into the HKU hive as MRU_PARSE!'
        if ' ' in sys.argv[1]:
            subprocess.call("reg LOAD HKU\MRU_PARSE "+'"'+sys.argv[1]+'"')
        else:
            subprocess.call("reg LOAD HKU\MRU_PARSE "+sys.argv[1])
    except IndexError:
        sys.exit('\nPlease supply the path to the NTUSER.dat file!')
    return

def reg_mru_framework(key):
    mapped_network = {}
    xref_lwt_ext = {}
    files_and_times=[]
    special = ['Map Network Drive MRU']
    #Check the ntuser.dat for the presence of the RecentDocs key.
    print '\n[+] Checking for RecentDocs key'
    try:
        subkeys = get_recentdoc_subkeys(key)
        print '\n[+] Key found!'
    except:
        print '\n[-] \Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs was not found!'
        sys.exit()
    try:
        try:
            given_key = _winreg.OpenKey(_winreg.HKEY_USERS, key)
            value_name_data = query_MRU(given_key)
            if key.split('\\').pop(-1) in special:
                for mapped_key, mapped_value in value_name_data.iteritems():
                    if mapped_key == 'MRUList':
                        mapped_network['MRUListEx'] = mapped_value.decode('hex')
                    else:
                        mapped_network[mapped_key] = mapped_value
                order = [mapped_network['MRUListEx'][i:i+1] for i in range(0, len(mapped_network['MRUListEx']), 1)]
            else:
                order = parse_MRU(value_name_data['MRUListEx'])
                
            file_names = get_order(value_name_data, order)
            working_mru_list= menu_display(file_names)
            MRU_list =  menu_display(file_names)
            mru_file = working_mru_list.pop(0).split('] ').pop(1)
            last_mod = _winreg.QueryInfoKey(given_key)[2]
            files_and_times.append({mru_file:str(query_last_write_time(last_mod))})
        except WindowsError:
            pass

        for sub_key in subkeys:
            extension_key = key+'\\'+sub_key
            try:
                ext_key = _winreg.OpenKey(_winreg.HKEY_USERS, extension_key)
                value = query_MRU(ext_key)
                xref_lwt_ext[sub_key]=value
            except WindowsError:
                pass
            
        for k,v in xref_lwt_ext.iteritems():
            last_mod, mru_file = parse_subkeys(key+'\\'+k, v)
            files_and_times.append({mru_file:str(query_last_write_time(last_mod))})


        return MRU_list, files_and_times
    except UnboundLocalError:
        sys.exit('Program needs to be executed from adminstrator prompt!')



def get_recentdoc_subkeys(key):
    sub_key_list =[]
    key = _winreg.OpenKey(_winreg.HKEY_USERS, key)
    try:
        sub_keys,values,time = _winreg.QueryInfoKey(key)
        for i in xrange(sub_keys):
            subkey = _winreg.EnumKey(key, i)
            sub_key_list.append(subkey)
    except WindowsError:
        print'here'
        pass
    return sub_key_list



def query_MRU(key):
    #queries a key using the QueryInfoKey and EnumValue.  I am interested in the MRUListEx key and value
    value_name_data = {}
    for i in xrange(0, _winreg.QueryInfoKey(key)[1]):
    	name, data, type = _winreg.EnumValue(key, i)
        value_name_data[name]=data.encode('hex')
    return value_name_data




def parse_MRU(MRUList):
    # I pass the value of the MRUListEx to this function which sorts the reading 0 and every 4th byte as hex
    # to get the order of file access from the MRUListEx.  Order is retunred in a list
    # i.e. [01, 23, 02, 45, 34]
    
    order_of_access = []
    data = [MRUList[i:i+2] for i in range(0, len(MRUList), 2)]
    try:
    	for x in range(0, len(data), 4):
        	order_of_access.append(int(data[x], 16))
    except ValueError:
        pass
    return order_of_access



def get_order(value_name_data, order):
    # value_name_data is dict set up like {value:data}
    # value = subkey value (i.e. 23) and data = hex string
    # iter through the list order and if i in order == key in value_name_data
    # append the hex string to file_names list
    # bascially build a hex sting list of file access order 
    file_names = []
    for x in order:
    	for k,v in value_name_data.iteritems():
            if k == str(x):
                data = [v[i:i+2] for i in range(0, len(v), 2)]
            	f_name = v.split('000000').pop(0)
                if len(f_name) % 2 !=0:
                    f_name = f_name+'0'
                file_names.append(f_name)
    return file_names



def menu_display(file_names):
    # build a menu of the hex string list
    # decode the hex strings to ascii
    counter = 1
    menu = []
    for i in file_names:
    	i=i.replace('00','')
    	menu_item = '[%d] %s' % (counter, i.decode('hex'))
    	menu.append(menu_item)
    	counter +=1
    return menu



def query_last_write_time(time_stamp):
    last_mod = datetime.utcfromtimestamp(float(time_stamp)*1e-7 - 11644473600)
    return last_mod



def parse_subkeys(key, value):
    key_order = parse_MRU(value['MRUListEx'])
    key_file_names = get_order(value, key_order)
    key_MRU_list = menu_display(key_file_names)
    mru_file = key_MRU_list.pop(0).split('] ').pop(1)
    open_key = _winreg.OpenKey(_winreg.HKEY_USERS, key)
    last_mod = _winreg.QueryInfoKey(open_key)[2]
    return last_mod, mru_file



def sort_final_list(menu_list):
    counter = 1
    menu = []
    for i in menu_list:
    	menu_item = '[%d] %s' % (counter, i)
    	check_format = len(menu_item)
    	menu.append(menu_item)
    	counter +=1
    return menu



def sort(mru_list, mru_timestamps):

    new_mru_list = []
    mru_with_stamps = []
    length_list = []

    #copy the keys (filenames)  into a new list to find the longest file name 
    for x in mru_timestamps:
        for k,v in x.iteritems():
            length_list.append(k)

    #establish proper spacing for format       
    format_space = max(length_list, key=len)
    length = len(format_space)+30

    #xref the mru_list with the mru_timestamps to build a timeline
    for file_name in mru_list:
        new_name = file_name.split('] ').pop(1)
        new_mru_list.append(new_name)
    for file_time in mru_timestamps:
        for k,v in file_time.iteritems():
            if k in new_mru_list:
                index = new_mru_list.index(k)
                length_k = len(k)
                length_v = len(v)
                if length_k + length_v <= length:
                    current_length = length_k + length_v
                    num_space = length - current_length-3
                    space = ' '*num_space
                    new_mru_list[index] = k+space+v
                else:
                    new_mru_list[index] = k+'      '+v
    return '\n'.join(sort_final_list(new_mru_list))



def write_to_file(location, final_list):
    try:
        with open(location, 'a') as f:
            f.write('MRU_PARSE OUTPUT\n')
            f.write('All Times are in UTC format\n\n\n')
            f.write(final_list)
            f.close()
    except IOError:
        unload()
        sys.exit('\n[X] IOError, check your file output path!\n[X] Printing output in terminal\n\n'+final_list)



def unload():
    print '\n[+] Unloading the '+sys.argv[1]+' from the registry!'
    subprocess.call("reg UNLOAD HKU\MRU_PARSE")
    return
    

if __name__ == '__main__':
    try:
        help_cmds = ['-help', '-h', 'help', '?']
        if sys.argv[1] in help_cmds:
            help_cmd()
        else:
            load()
            key = r'MRU_PARSE\Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs'
            mru_list, mru_timestamps = reg_mru_framework(key)
            final_list = sort(mru_list, mru_timestamps)
            try:
                if len(sys.argv[2]) != 0:
                    write_to_file(sys.argv[2], final_list)
                    print '\n[+] Complete...written to '+sys.argv[2]
            except IndexError:
                print '\n'+final_list
            unload()
            
    except KeyboardInterrupt:
        sys.exit('[x] Shuttng Down!')



