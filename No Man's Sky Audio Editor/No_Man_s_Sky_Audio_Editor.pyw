import cPickle as pickle
import Tkinter, ttk, tkFileDialog, tkMessageBox, os, shutil, thread
from subprocess import call, Popen, PIPE

def set_window_title(str):
    global window_title
    window_title = str
    root.title(window_title)

def new_project():

    global edit_list
    edit_list = []
    settings['project'] = ''
    save_settings()
    for item in filetree.get_children():
        filetree.delete(item)
    populate_tree()
    set_window_title('NMS Audio Editor')

def open_project():
    
    file_name = tkFileDialog.askopenfilename(initialdir = "/", title = "Open NMSAE Project", filetypes = [("NMSAE project file",("*.nmsae"))])
    if file_name == '':
        return
    load_project(file_name)
    settings['project'] = file_name
    save_settings()

def load_project(path):

    if path == '':
        return
    global edit_list
    f = open(path)
    edit_list = pickle.load(f)
    f.close()
    for item in edit_list:
        filetree.set(item['id'], 'new', item['new'])
    set_window_title('NMS Audio Editor | ' + os.path.basename(path))

def save_project():

    save_file = tkFileDialog.asksaveasfilename(initialdir = "/", title = "Save Project as...", defaultextension=".nmsae", filetypes = (("NMSAE project file","*.nmsae"), ("All Files", "*.*")))
    if save_file == '':
        return
    f = open(save_file, 'w')
    pickle.dump(edit_list, f)
    f.close()
    settings['project'] = save_file
    save_settings()
    set_window_title('NMS Audio Editor | ' + os.path.basename(save_file))

def populate_tree():
    for item in flist:
        full_path = item['in_path']
        parent = full_path.split('\\')
        parent.insert(0, '')
        child = parent.pop()
        for i in range(1, len(parent)):
            ppath = '\\'.join(parent[0:i])
            cpath = '\\'.join(parent[0:i + 1])
            if cpath not in filetree.get_children(ppath):
                filetree.insert(ppath, 'end', cpath, text=parent[i])
        parent = '\\'.join(parent)
        filetree.insert(parent, 'end', item['id'], text=child, values=(item['id'], '', item['in_bnk']))
    
def check_root_audio():
    """Checks if 'AUDIO' folder exists in root of drive. If so, display error and exit"""
    
    if os.path.isdir('\\AUDIO'):
        try:
            os.rmdir('\\AUDIO')
        except:
            if tkMessageBox.askokcancel('Warning', '''You have a non-empty folder called "AUDIO"
on the root level of your drive. This
program requires the use of that folder
in order to run properly. Plus, it is
designed to delete the "AUDIO" folder when
it's done. I don't want to delete your files accidentally.
Please move, rename, or delete that folder and try again.

Press OK to quit. Press Cancel to delete "AUDIO" and continue.'''):
                exit()
            else:
                shutil.rmtree('\\AUDIO')
                return

def find_dict(list, key, value):
    """Searches for dictionary in list with matching key:value, returns first match

    :param list: list of dictionaries
    :param key: key to match
    :param value: value to match
    """
    for item in list:
        if key in item:
            if item[key] == value:
                return item
    return

def del_file(path):
    """Deletes a file

    :param path: path to file to delete
    """
    if os.path.isfile(path):
        os.remove(path)
    return

def copy_wav(path):
    """Copies wav file to Wwise conversion input path
    
    :param path: path of input wav
    """
    del_file(wav_in)
    shutil.copyfile(path, wav_in)
    return

def convert_mp3_to_wav(path):
    """Converts and moves mp3 to Wwise conversion input path

    :param path: path of mp3
    """
    del_file(wav_in)
    call([paths['ffmpeg'], '-i', path, wav_in])
    return

def convert_wav_to_wem():
    """Converts wav_in to wem_out"""
    
    call([paths['wwise'], paths['template'], '-ConvertExternalSources', 'Windows'])
    del_file(wav_in)
    return

def pack_mod(output):
    """Uses psarc.exe to package AUDIO folder in root, then deletes AUDIO
    
    :param output: path to output PAK file
    """
    progress_text.set('Packing Mod...')
    out_file = '--output=' + output
    call([paths['psarc'], 'create', '-y', '\\AUDIO', out_file])
    shutil.rmtree('\\AUDIO')
    progress_text.set('Done')

def replace_selected_with(*args):
    """Prompts user for file to replace selected items with and adds to edit_list"""
    
    selected = filetree.selection()
    if len(selected) == 0:
        return
    file_name = tkFileDialog.askopenfilename(initialdir = "/", title = "Replace selected with...", filetypes = [("Audio files",("*.mp3","*.wav"))])
    if file_name == '':
        return
    for item in selected:
        if len(filetree.get_children(item)) == 0:
            item_id = filetree.set(item, 'id')
            filetree.set(item, 'new', file_name)
            in_list = find_dict(edit_list, 'id', item_id)
            if in_list == None:
                entry = find_dict(flist, 'id', item_id)
                entry['new'] = file_name
                edit_list.append(entry)
            else:
                in_list['new'] = file_name
    set_window_title(window_title + ' *')

def set_widget_state(list, state):
    """Sets state of multible ttk widgets

    :param list: list of ttk widget objects
    :param state: string value ttk widget state
    """

    for item in list:
        item.state([state])

def find_dups(edit_list):
    """Sorts out entries of edit_list into a dictionary of lists each with same replacemend file

    Returns this dict
    :param edit_list: edit_list
    """
    return_dict = {}
    for item in edit_list:
        new = item['new']
        if new not in return_dict:
            return_dict[new] = [item]
        else:
            return_dict[new].append(item)
    return return_dict

def build_mod(*args):
    """Moves, converts renames and packs replacement files using edit_list as input"""

    global edit_list
    global tmp_list
    build_button.state(['disabled'])
    tmp_list = []
    for item in edit_list:
        if 'NMS_AUDIO_PERSISTENT' in item['in_bnk']:
            result = tkMessageBox.askyesnocancel('Warning', '''You have selected to change a file that resides inside of the soundbank NMS_AUDIO_PERSISTENT.
Repacking this soundbank will likely take a LONG time (5 to 10+ min).
Are you sure your mod is ready?

"Yes" to continue build
"No" to build without NMS_AUDIO_PERSISTENT
"Cancel" to halt build''')
            if result == None:
                build_button.state(['!disabled'])
                return
            elif result == False:
                for thing in edit_list:
                    if 'NMS_AUDIO_PERSISTENT' in thing['in_bnk']:
                        tmp_list.append(thing)
                        edit_list.remove(thing)
            break
    if edit_list == []:
        edit_list += tmp_list
        build_button.state(['!disabled'])
        return
    save_file = tkFileDialog.asksaveasfilename(initialdir = "/", title = "Save Mod as...", defaultextension=".pak", filetypes = (("NMS Packed Mod","*.pak"), ("All Files", "*.*")))
    if save_file == '':
        edit_list += tmp_list
        build_button.state(['!disabled'])
        return
    shutil.copytree(paths['AUDIO'], '\\AUDIO')
    dups = find_dups(edit_list)
    bnks = check_bnk(edit_list)
    progress_value.set(0)
    thread.start_new_thread(convert, (dups, bnks, save_file))

def prepare_bnks(bnks):
    """Extracts necessary BNKs with soundMod

    :param bnks: bnks list
    """
    
    tmp = os.getcwd()
    os.chdir(paths['sound_mod'])
    paths['oggdec'] = '..\\oggdec.exe'
    progress_text.set('Preparing BNKs...')
    text = ''
    for bnk in bnks:
        text += os.path.basename(bnk).replace('.BNK','') + '\n'
    f = open('sound_banks.txt', 'w')
    f.write(text)
    f.close()
    process = Popen('SoundFileEditor.exe', stdout=PIPE)
    for i in range(0,3):
        process.stdout.readline()
    call(['taskkill', '/f', '/im', 'SoundFileEditor.exe'])
    os.chdir(tmp)
    paths['oggdec'] = 'files\\oggdec.exe'

def ready_bnk(item):
    """Copies wem_out to appropriate soundMod folder and adds its name to data.txt

    :param item: item that is being replaced in bnk
    """

    item_id = item['id']
    new_name = item_id + '_r.wem'
    folder_path = paths['sound_mod'] + os.path.basename(item['in_bnk']).replace('.BNK', '')
    shutil.copyfile(wem_out, os.path.join(folder_path, new_name))
    old_line = item_id + '=#'
    new_line = item_id + '=' + new_name + '#'
    replace_in_file(os.path.join(folder_path, 'data.txt'), old_line, new_line)

def pack_bnks():
    tmp = os.getcwd()
    os.chdir(paths['sound_mod'])
    paths['oggdec'] = '..\\oggdec.exe'
    progress_text.set('Packing BNKs...')
    to_delete = ['Output']
    f = open('sound_banks.txt', 'r')
    for line in f:
        to_delete.append(line.replace('\n', ''))
    f.close()
    process = Popen('SoundFileEditor.exe', stdout=PIPE)
    for i in range(0,3):
        process.stdout.readline()
    call(['taskkill', '/f', '/im', 'SoundFileEditor.exe'])
    for out in os.listdir('Output'):
        out = os.path.join('Output', out)
        out_fix = out.replace('bnk', 'BNK')
        os.rename(out, out_fix)
        shutil.copy(out_fix, '\\AUDIO')
    for folder in to_delete:
        shutil.rmtree(folder)
    os.chdir(tmp)
    paths['oggdec'] = 'files\\oggdec.exe'

def convert(dups, bnks, save_file):
    global edit_list
    if len(bnks) > 0:
        prepare_bnks(bnks)
    i = 0
    for new in dups:
        i += 1
        progress_text.set('Converting {} of {}'.format(i, len(dups)))
        if not audio_file(new):
            continue
        convert_wav_to_wem()
        for item in dups[new]:
            if item['in_bnk'] == '':
                end = '\\' + item['out_path']
                shutil.copyfile(wem_out, end)
            else:
                ready_bnk(item)
        progress_value.set(100.0 * i / len(dups))
    
    if len(bnks) > 0:
        pack_bnks()
    pack_mod(save_file)
    edit_list += tmp_list    
    build_button.state(['!disabled'])

def audio_file(path):
    """Check file type and performs appropriate action

    :param path: path to file
    """

    if path.lower().endswith('mp3'):
        convert_mp3_to_wav(path)
    elif path.lower().endswith('wav'):
        copy_wav(path)
    else:
        return False
    return True

def clear_selected(*args):
    """Remove selected items from edit_list"""

    selected = filetree.selection()
    if len(selected) == 0:
        return
    for item in selected:
        if len(filetree.get_children(item)) == 0:
            item_id = filetree.set(item, 'id')
            filetree.set(item, 'new', '')
            in_list = find_dict(edit_list, 'id', item_id)
            if in_list != None:
                edit_list.remove(in_list)     

def save_settings():
    f = open(paths['settings'], 'w')
    pickle.dump(settings, f)
    f.close()

def select_audio_dump(*args):
    """Allow user to set the location of Audio dump for use in playback"""

    settings['dump_path'] = tkFileDialog.askdirectory(parent=root, initialdir='.', title='Please select a the location of the audio dump')
    save_settings()

def play_audio(*args):
    """When nothing is playing: Plays first selected audio file
    
    When something is playing:
    Stops playing if same file selected
    Stops old and starts new if different file selected
    """
    global playing
    selection = filetree.selection()
    for item in selection:
        if len(filetree.get_children(item)) == 0:
            item_id = filetree.set(item, 'id')
            item = find_dict(flist, 'id', item_id)
            file_path = os.path.join(settings['dump_path'], item['in_path'])
            if playing is not None:
                playing[0].terminate()
                if playing[1] == file_path:
                    playing = None
                    return
            if os.path.isfile(file_path):
                playing = [Popen([paths['oggdec'], '-p', "{}".format(file_path)]), file_path]
            else:
                tkMessageBox.showerror('Error', '''Cannot find audio file!
Are you sure you selected the correct location of the Audio Dump?''')
            return

def check_bnk(edit_list):
    """Sorts out entries of edit_list into a dictionary of lists each with same BNK

    Returns this dict
    :param edit_list: edit_list
    """
    return_dict = {}
    for item in edit_list:
        bnk = item['in_bnk']
        if bnk == '':
            continue
        if bnk not in return_dict:
            return_dict[bnk] = [item]
        else:
            return_dict[bnk].append(item)
    return return_dict

def replace_in_file(file_path, old, new):
    """Replaces all instances of old with new in file

    :param file_paht: path to file
    :param old: search string
    :param new: replace string
    """

    f = open(file_path, 'r')
    text = f.read()
    f.close()    
    text = text.replace(old, new)
    f = open(file_path, 'w')
    f.write(text)
    f.close()

#set constants
paths = {'flist':'files\\filelist.pkl',
        'settings':'files\\settings.pkl',
        'wwise':os.getenv('WWISEROOT') + 'Authoring\\Win32\\Release\\bin\\WwiseCLI.exe',
        'ffmpeg':'files\\ffmpeg.exe',
        'psarc':'files\\psarc.exe',
        'oggdec':'files\\oggdec.exe',
        'template':'files\\Wwise_Template\\Template.wproj',
        'sound_mod':'files\\soundMod\\',
        'AUDIO':'files\\AUDIO'
        }
wav_in = 'files\\Wwise_Template\\Input\\in.wav'
wem_out = 'files\\Wwise_Template\\Output\\out.wem'
edit_list = []
playing = None

#load settings
f = open(paths['settings'])
settings = pickle.load(f)
f.close()


#load file/directory structure
f = open(paths['flist'])
flist = pickle.load(f)
f.close()

#construct the UI
#make a window
root = Tkinter.Tk()
set_window_title('NMS Audio Editor')
root.geometry('{}x{}'.format(int(root.winfo_screenwidth()*.85), int(root.winfo_screenheight()*.75)))
root.rowconfigure(0, weight=1)
root.columnconfigure(0, weight=1)

menu_bar = Tkinter.Menu(root)

file_menu = Tkinter.Menu(menu_bar, tearoff=0)
file_menu.add_command(label='New Project', command=new_project)
file_menu.add_command(label='Open Project', command=open_project)
file_menu.add_command(label='Save Project', command=save_project)
menu_bar.add_cascade(label='File', menu=file_menu)

root.config(menu=menu_bar)

#top level frame
sup = ttk.Frame(root, padding=(5,5,5,5))
sup.grid(column=0, row=0, sticky='NSEW')
sup.rowconfigure(0, weight=1)
sup.columnconfigure(0, weight=1)

filetree_frame = ttk.Frame(sup)
filetree_frame.grid(column=0, row=0, sticky="NSEW")
filetree_frame.rowconfigure(0, weight=1)
filetree_frame.columnconfigure(0, weight=1)


filetree = ttk.Treeview(filetree_frame, columns=('id','new','in_bnk'))
filetree.column('#0', width=300, minwidth=150, stretch=0)
filetree.column('id', width=75, minwidth=75, stretch=0)
filetree.column('new', width=150, minwidth=150, stretch=0)
filetree.heading('#0', text='Audio File', anchor='w')
filetree.heading('id', text='ID', anchor='w')
filetree.heading('new', text='Replace With', anchor='w')
filetree.heading('in_bnk', text='In BNK', anchor='w')
filetree.bind('<Double-1>', replace_selected_with)
filetree.bind('p', play_audio)
filescroll = ttk.Scrollbar(filetree_frame, orient=Tkinter.VERTICAL, command=filetree.yview)
filetree.configure(yscrollcommand=filescroll.set)
filetree.grid(column=0, row=0, sticky="NSEW")
filescroll.grid(column=1, row=0, sticky='NSEW', padx=(0, 0), pady=(5, 5))

populate_tree()

right_frame = ttk.Frame(sup)
right_frame.grid(column=1, row=0, sticky='NSEW')
right_frame.rowconfigure(0, weight=1)

instructions_frame = ttk.Frame(right_frame)
instructions_frame.grid(column=0, row=0, sticky="NSEW")

#buttons ahead
button_frame = ttk.Frame(right_frame)
button_frame.grid(column=0, row=1, sticky="NSEW")

folder_button = ttk.Button(button_frame, text='Select Audio Dump Folder', command=select_audio_dump)
play_button = ttk.Button(button_frame, text='Play/Stop selected Audio', command=play_audio)
repl_button = ttk.Button(button_frame, text='Replace Selected With...', command=replace_selected_with)
clear_button = ttk.Button(button_frame, text='Clear Selected', command=clear_selected)
build_button = ttk.Button(button_frame, text='Build Mod', command=build_mod)
button_list = [folder_button, play_button, repl_button, clear_button, build_button]
for i in range(0, len(button_list)):
    button_list[i].grid(column=0, row=(i), sticky='NSEW')

progress_frame = ttk.Frame(right_frame)
progress_frame.grid(column=0, row=2, sticky="NSEW")
progress_frame.columnconfigure(0, weight=1)

progress_text = Tkinter.StringVar()
progress_text.set('')
progress_label = ttk.Label(progress_frame, textvariable=progress_text)
progress_label.grid(column=0, row=0, sticky="NSEW")

progress_value = Tkinter.IntVar()
progress_value.set(0)
progress_bar = ttk.Progressbar(progress_frame, orient=Tkinter.HORIZONTAL, mode='determinate', variable=progress_value)
progress_bar.grid(column=0, row=1, sticky="NSEW")

#throw on a sizegrip
ttk.Sizegrip(root).grid(column=0, row=1, sticky='SE')





check_root_audio()
load_project(settings['project'])


#LOOP IT!!!
root.mainloop()