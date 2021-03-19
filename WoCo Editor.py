import tkinter as tk
from tkinter import ttk
from tkinter import font, colorchooser, filedialog, messagebox
import os
import self as self
import win32api
import wikipedia
from tkinter import *
from win32com.client import Dispatch

main_application = tk.Tk()
main_application.geometry('1200x800')
main_application.title('WoCo Editor')
main_application.wm_iconbitmap('icon.ico') # header small icon


######### main menu #######

main_menu = tk.Menu()

file = tk.Menu(main_menu, tearoff=False) #tearoff seperate menu
#file icons
new_icon = tk.PhotoImage(file='icons/new.png')  #photo image class to use icons#
open_icon = tk.PhotoImage(file='icons/open.png')
save_icon = tk.PhotoImage(file='icons/save.png')
save_as_icon = tk.PhotoImage(file='icons/save_as.png')
print_icon = tk.PhotoImage(file='icons/printer.png')
exit_icon = tk.PhotoImage(file='icons/exit.png')


edit = tk.Menu(main_menu, tearoff=False)
#edit icons
copy_icon = tk.PhotoImage(file='icons/copy.png')
paste_icon = tk.PhotoImage(file='icons/paste.png')
cut_icon = tk.PhotoImage(file='icons/cut.png')
clear_all_icon = tk.PhotoImage(file='icons/clear_all.png')
find_icon = tk.PhotoImage(file='icons/find.png')


view = tk.Menu(main_menu, tearoff=False)
#view icons
tool_bar_icon = tk.PhotoImage(file='icons/tool_bar.png')
status_bar_icon = tk.PhotoImage(file='icons/status_bar.png')


color_theme = tk.Menu(main_menu, tearoff=False)
#color theme
light_default_icon = tk.PhotoImage(file='icons/light_default.png')
light_plus_icon = tk.PhotoImage(file='icons/light_plus.png')
dark_icon = tk.PhotoImage(file='icons/dark.png')
red_icon = tk.PhotoImage(file='icons/red.png')
night_blue_icon = tk.PhotoImage(file='icons/night_blue.png')
monokai_icon = tk.PhotoImage(file='icons/monokai.png')

theme_choice = tk.StringVar() #variable to store color selection value#
color_icons = (light_default_icon,light_plus_icon,dark_icon,red_icon,night_blue_icon,monokai_icon) #list or tuple#
#first for text second value for background#
color_dict = {
    'Light Default':('#000000','#ffffff'),
'Light Plus':('#474747','#e0e0e0'),
'Dark':('#c4c4c4','#2d2d2d'),
'Red':('#2d2d2d','#ffe8e8'),
'Monokai':('#d3b774','#474747'),
'Night Blue':('#ededed','#6b9dc2')
}

# help
help = tk.Menu(main_menu, tearoff=False)
help_icon = tk.PhotoImage(file='icons/help.png')



# redo and undo
redo = tk.Menu(main_menu, tearoff=False)
undo_icon = tk.PhotoImage(file='icons/undo.png')
redo_icon = tk.PhotoImage(file='icons/redo.png')

#cascade to show menu
main_menu.add_cascade(label='File', menu=file)
main_menu.add_cascade(label='Edit', menu=edit)
main_menu.add_cascade(label='View', menu=view)
main_menu.add_cascade(label='Color Theme', menu=color_theme)
main_menu.add_cascade(label='About', menu=help)
main_menu.add_cascade(label='Redo', menu=redo)

######### main menu end #######


######### toolbar #######

tool_bar = ttk.Label(main_application)
tool_bar.pack(side=tk.TOP, fill=tk.X) #to set tool bar on top and fill horizontally (use x to fill horizontally and y to vertically and both for both)

#font family box
font_tuple = tk.font.families() #all font families
font_family = tk.StringVar() #value of font
font_box = ttk.Combobox(tool_bar, width=30, textvariable=font_family, state='readonly') #read only because user has to select font only
font_box['values'] = font_tuple
font_box.current(font_tuple.index('Arial')) # use current to set default font family
font_box.grid(row=0, column=0, padx=5) # use grid to show tool bar

#size box
size_var = tk.IntVar()
font_size = ttk.Combobox(tool_bar, width=14, textvariable = size_var, state='readonly')
font_size['values'] = tuple(range(8,80)) # font size range 8 to 80 and difference of 2
font_size.current(4) # default value will be 8+4 = 12
font_size.grid(row=0, column=1, padx=5)

#bold button
bold_icon = tk.PhotoImage(file='icons/bold.png')
bold_btn = ttk.Button(tool_bar, image=bold_icon) # first parameter tell where to place button
bold_btn.grid(row=0, column=2, padx=5)

#italic button
italic_icon = tk.PhotoImage(file='icons/italic.png')
italic_btn = ttk.Button(tool_bar, image=italic_icon) # first parameter tell where to place button
italic_btn.grid(row=0, column=3, padx=5)

#underline button
underline_icon = tk.PhotoImage(file='icons/underline.png')
underline_btn = ttk.Button(tool_bar, image=underline_icon) # first parameter tell where to place button
underline_btn.grid(row=0, column=4, padx=5)

#font color button
font_color_icon = tk.PhotoImage(file='icons/font_color.png')
font_color_btn = ttk.Button(tool_bar, image=font_color_icon) # first parameter tell where to place button
font_color_btn.grid(row=0, column=5, padx=5)

#align left
align_left_icon = tk.PhotoImage(file='icons/align_left.png')
align_left_btn = ttk.Button(tool_bar, image=align_left_icon) # first parameter tell where to place button
align_left_btn.grid(row=0, column=6, padx=5)

#align center
align_center_icon = tk.PhotoImage(file='icons/align_center.png')
align_center_btn = ttk.Button(tool_bar, image=align_center_icon) # first parameter tell where to place button
align_center_btn.grid(row=0, column=7, padx=5)

#align right
align_right_icon = tk.PhotoImage(file='icons/align_right.png')
align_right_btn = ttk.Button(tool_bar, image=align_right_icon) # first parameter tell where to place button
align_right_btn.grid(row=0, column=8, padx=5)

# clear format
clear_format_icon = tk.PhotoImage(file='icons/clear_format.png')
clear_format_btn = ttk.Button(tool_bar, image=clear_format_icon) # first parameter tell where to place button
clear_format_btn.grid(row=0, column=9, padx=5)

# speak button
speak_icon = tk.PhotoImage(file='icons/speak.png')
speak_btn = ttk.Button(tool_bar, image=speak_icon) # first parameter tell where to place button
speak_btn.grid(row=0, column=10, padx=5)

# wikipedia button
wikipedia_icon = tk.PhotoImage(file='icons/wikipedia.png')
wikipedia_btn = ttk.Button(tool_bar, image=wikipedia_icon) # first parameter tell where to place button
wikipedia_btn.grid(row=0, column=11, padx=5)


######### toolbar end #######

######### text editor #######

text_editor = tk.Text(main_application, undo=True)
text_editor.config(wrap='word', relief=tk.FLAT)

scroll_bar = tk.Scrollbar(main_application) #to make scroll bar
text_editor.focus_set() # to make cursor blink
scroll_bar.pack(side=tk.RIGHT, fill=tk.Y)
text_editor.pack(fill=tk.BOTH, expand=True)
scroll_bar.config(command=text_editor.yview) # to tell scroll bar is for text
text_editor.config(yscrollcommand=scroll_bar.set) # to set scroll bar vertically

# font family and font size functionality
current_font_family = 'Arial'
current_font_size = 12

def change_font(event=None):
    global current_font_family
    current_font_family = font_family.get()
    text_editor.configure(font=(current_font_family, current_font_size))
    

def change_fontsize(event=None):
    global current_font_size
    current_font_size = size_var.get()
    text_editor.configure(font=(current_font_family, current_font_size))

font_box.bind("<<ComboboxSelected>>", change_font)
font_size.bind("<<ComboboxSelected>>", change_fontsize)

# bold buttons functionality
def change_bold():
    text_property = tk.font.Font(font=text_editor['font'])
    if text_property.actual()['weight'] == 'normal':
        text_editor.configure(font=(current_font_family, current_font_size, 'bold'))
    if text_property.actual()['weight'] == 'bold':
        text_editor.configure(font=(current_font_family, current_font_size, 'normal'))

bold_btn.configure(command=change_bold) # giving command function to work


# italic button functionality
def change_italic():
    text_property = tk.font.Font(font=text_editor['font'])
    if text_property.actual()['slant'] == 'roman':
        text_editor.configure(font=(current_font_family, current_font_size, 'italic'))
    if text_property.actual()['slant'] == 'italic':
        text_editor.configure(font=(current_font_family, current_font_size, 'roman'))

italic_btn.configure(command=change_italic) # giving command function to work

# underline button functionality

def change_underline():
    text_property = tk.font.Font(font=text_editor['font'])
    if text_property.actual()['underline'] == 0:
        text_editor.configure(font=(current_font_family, current_font_size, 'underline'))
    if text_property.actual()['underline'] == 1:
        text_editor.configure(font=(current_font_family, current_font_size, 'normal'))

underline_btn.configure(command=change_underline) # giving command function to work

# change font color functionality
def change_font_color():
    color_var = tk.colorchooser.askcolor()
    text_editor.configure(fg=color_var[1]) # fg is shortcut for foreground text color


font_color_btn.configure(command=change_font_color)

# align left text functionality
def align_left():
    text_content = text_editor.get(1.0, 'end') # get everything from start to finish
    text_editor.tag_config('left', justify=tk.LEFT) #configure to justify left
    text_editor.delete(1.0, tk.END) # first delete and then paste on specified location
    text_editor.insert(tk.INSERT, text_content, 'left') # fisrt argument to insert and second for what to insert and last where to insert

align_left_btn.configure(command=align_left)



# align center text functionality
def align_center():
    text_content = text_editor.get(1.0, 'end') # get everything from start to finish
    text_editor.tag_config('center', justify=tk.CENTER) #configure to justify left
    text_editor.delete(1.0, tk.END) # first delete and then paste on specified location
    text_editor.insert(tk.INSERT, text_content, 'center') # fisrt argument to insert and second for what to insert and last where to insert

align_center_btn.configure(command=align_center)



# align right text functionality
def align_right():
    text_content = text_editor.get(1.0, 'end') # get everything from start to finish
    text_editor.tag_config('right', justify=tk.RIGHT) #configure to justify left
    text_editor.delete(1.0, tk.END) # first delete and then paste on specified location
    text_editor.insert(tk.INSERT, text_content, 'right') # fisrt argument to insert and second for what to insert and last where to insert

align_right_btn.configure(command=align_right)



# format clear functionality
def format_clear():
    text_property = tk.font.Font(font=text_editor['font'])
    if text_property.actual()['weight'] == 'bold':
        text_editor.configure(font=(current_font_family, current_font_size, 'normal'))
    if text_property.actual()['slant'] == 'italic':
        text_editor.configure(font=(current_font_family, current_font_size, 'roman'))
    if text_property.actual()['underline'] == 1:
        text_editor.configure(font=(current_font_family, current_font_size, 'normal'))

clear_format_btn.configure(command=format_clear)

# speak button functionality
def open_speak():
    # speak = Dispatch("SAPI.SpVoice")
    speak = Dispatch("SAPI.SpVoice")

    def talk(x):
        speak.Speak(x.get('1.0', tk.END))

    root = tk.Tk()
    root.title("WoCoPad Speak Assistant")
    root.geometry("600x400")
    root.wm_iconbitmap('icon.ico')
    label = tk.Label(root, text="Enter Text Below To Speak Aloud").pack()
    entry = tk.Text(root)
    entry.insert("1.0", "How can i help you?")
    button = tk.Button(root, text="Press Me To Speak", command=lambda: talk(entry))
    button.pack()
    entry.pack()
    entry.focus()
    root.mainloop()
speak_btn.configure(command=open_speak)

# Wikipedia button functionality
def open_wiki():
    def get_me():
        entry_value = entry.get()
        answer.delete(1.0, END)
        try:
            answer_value = wikipedia.summary(entry_value)
        except:
            answer.insert(INSERT, "Please Check Your Internet Connection OR Input")

        answer.insert(INSERT, answer_value)
    root = Tk()
    root.title("Wikipedia Assistant")
    root.geometry('600x400')
    topframe = Frame(root)
    entry = Entry(topframe)
    entry.pack()
    button = Button(topframe, text="search", command=get_me)
    button.pack()
    topframe.pack(side=TOP)
    bottomframe = Frame(root)
    scroll = Scrollbar(bottomframe)
    scroll.pack(side=RIGHT, fill=Y)
    answer = Text(bottomframe, yscrollcommand=scroll.set, wrap=WORD, width=600, height=400)
    scroll.config(command=answer.yview)
    answer.pack()
    bottomframe.pack()
    root.mainloop()
wikipedia_btn.configure(command=open_wiki)


text_editor.configure(font=('Arial', 12))
######### text editor end #######

######### status bar #######

status_bar = ttk.Label(main_application, text= 'Status Bar') #at start status bar show the text
status_bar.pack(side=tk.BOTTOM) #to fix it at bottom (using pack for L,R,B,T) and (grid for row and coloumn)

text_changed = False # variable for exit menu to check whether text is modified or not

def changed(event=None):
    global text_changed
    if text_editor.edit_modified():
        text_changed = True
        words = len(text_editor.get(1.0, 'end-1c').split()) # to avoid counting new line use end-1c and split to seperate words and get to get content
        characters = len(text_editor.get(1.0, 'end-1c').replace(' ', '')) #replace function to avoid counting spaces as charater
        status_bar.config(text=f'Characters : {characters} Words : {words}')
    text_editor.edit_modified(False) # if text is not modify

text_editor.bind('<<Modified>>', changed)

######### status bar end #######



######### main menu functionality #######

## global variable
url = ''

## new functionality
def new_file(event=None):
    global url
    url = ''
    text_editor.delete(1.0, tk.END)

## open functionality

def open_file(event=None):
    global url #file dialog to ask user to choose file to be open                              # * is file name
    url = filedialog.askopenfilename(initialdir=os.getcwd(), title='Select File', filetypes=(('Text File', '*.txt'), ('All files', '*.*'))) # os to define current working directory
    try: # file not found error
        with open(url, 'r') as fr:
            text_editor.delete(1.0, tk.END)
            text_editor.insert(1.0, fr.read()) # file reader
    except FileNotFoundError:
        return
    except:
        return
    main_application.title(os.path.basename(url))

## save file

def save_file(event=None):
    global url
    try:
        if url:
            content = str(text_editor.get(1.0, tk.END))
            with open(url, 'w', encoding='utf-8') as fw:
                fw.write(content)
        else:
            url = filedialog.asksaveasfile(mode = 'w', defaultextension='.txt', filetypes=(('Text File', '*.txt'), ('All files', '*.*')))
            content2 = text_editor.get(1.0, tk.END)
            url.write(content2)
            url.close()
    except:
        return

## save as functionality
def save_as(event=None):
    global url
    try:
        content = text_editor.get(1.0, tk.END)
        url = filedialog.asksaveasfile(mode = 'w', defaultextension='.txt', filetypes=(('Text File', '*.txt'), ('All files', '*.*')))
        url.write(content)
        url.close()
    except:
        return

## Print file functionality
def print_file():
    file_to_print = filedialog.askopenfilename(initialdir=r'C:\Users', title='Select File', filetypes=(('Text File', '*.txt'), ('All files', '*.*')))
    if file_to_print:
        win32api.ShellExecute(0, "print", file_to_print, None, ".", 0)

## exit functionality

def exit_func(event=None):
    global url, text_changed
    try:
        if text_changed:
            mbox = messagebox.askyesnocancel('Warning', 'Do you want to save the file ?')
            if mbox is True:
                if url:
                    content = text_editor.get(1.0, tk.END)
                    with open(url, 'w', encoding='utf-8') as fw:
                        fw.write(content)
                        main_application.destroy()
                else:
                    content2 = str(text_editor.get(1.0, tk.END))
                    url = filedialog.asksaveasfile(mode = 'w', defaultextension='.txt', filetypes=(('Text File', '*.txt'), ('All files', '*.*')))
                    url.write(content2)
                    url.close()
                    main_application.destroy()
            elif mbox is False: # dont want to save file click no it closes message box
                main_application.destroy()
        else: # click yes to save file
            main_application.destroy()
    except:
        return



# find functionality
def find_func(event=None):
    def find(): # function to find
        word = find_input.get()
        text_editor.tag_remove('match', '1.0', tk.END)
        matches = 0
        if word:
            start_pos = '1.0'
            while True:
                start_pos = text_editor.search(word, start_pos, stopindex=tk.END)
                if not start_pos:
                    break
                end_pos = f'{start_pos}+{len(word)}c'
                text_editor.tag_add('match', start_pos, end_pos)
                matches += 1
                start_pos = end_pos
                text_editor.tag_config('match', foreground='red', background='yellow')

    def replace(): # function to replace find text
        word = find_input.get()
        replace_text = replace_input.get()
        content = text_editor.get(1.0, tk.END)
        new_content = content.replace(word, replace_text)
        text_editor.delete(1.0, tk.END)
        text_editor.insert(1.0, new_content)

    find_dialogue = tk.Toplevel() # to create pop up window
    find_dialogue.geometry('450x250+500+200') # to set window size
    find_dialogue.title('Find')
    find_dialogue.resizable(0, 0)

    ## frame
    find_frame = ttk.LabelFrame(find_dialogue, text='Find/Replace') # first argument tell where to create label frame
    find_frame.pack(pady=20)

    ## labels
    text_find_label = ttk.Label(find_frame, text='Find : ')
    text_replace_label = ttk.Label(find_frame, text='Replace')

    ## entry
    find_input = ttk.Entry(find_frame, width=30)
    replace_input = ttk.Entry(find_frame, width=30)

    ## button
    find_button = ttk.Button(find_frame, text='Find', command=find)
    replace_button = ttk.Button(find_frame, text='Replace', command=replace)

    ## label grid
    text_find_label.grid(row=0, column=0, padx=4, pady=4)
    text_replace_label.grid(row=1, column=0, padx=4, pady=4)

    ## entry grid
    find_input.grid(row=0, column=1, padx=4, pady=4)
    replace_input.grid(row=1, column=1, padx=4, pady=4)

    ## button grid
    find_button.grid(row=2, column=0, padx=8, pady=4)
    replace_button.grid(row=2, column=1, padx=8, pady=4)

    find_dialogue.mainloop()
# find functionality end









# help functionality
def about():
    description = messagebox.showinfo("About WoCoPad", "WoCoPad\nv1.0\nA Notepad made by Hamza Khan.")


# file commands
file.add_command(label='New', image=new_icon, compound=tk.LEFT, accelerator='Ctrl+N', command=new_file)#to set commands and use compound to stop overlay of text and image#
file.add_command(label='Open', image=open_icon, compound=tk.LEFT, accelerator='Ctrl+O', command=open_file)#accelerator to add text front of another text
file.add_command(label='Save', image=save_icon, compound=tk.LEFT, accelerator='Ctrl+S', command=save_file)
file.add_command(label='Save As', image=save_as_icon, compound=tk.LEFT, accelerator='Ctrl+Alt+S', command=save_as)
file.add_command(label='Print File', image=print_icon, compound=tk.LEFT, accelerator='Print A File', command=print_file)
file.add_command(label='Exit', image=exit_icon, compound=tk.LEFT, accelerator='Ctrl+Q', command=exit_func)

# edit commands
edit.add_command(label='Copy', image=copy_icon, compound=tk.LEFT, accelerator='Ctrl+C', command=lambda:text_editor.clipboard_append(self.text_editor.selection_get()))
edit.add_command(label='Paste', image=paste_icon, compound=tk.LEFT, accelerator='Ctrl+V',command=lambda:text_editor.event_generate("<Control v>"))
edit.add_command(label='Cut', image=cut_icon, compound=tk.LEFT, accelerator='Ctrl+X', command=lambda:text_editor.delete("sel.first", "sel.last"))
edit.add_command(label='Clear All', image=clear_all_icon, compound=tk.LEFT, accelerator='Ctrl+Alt+X', command=lambda:text_editor.delete(1.0, tk.END))
edit.add_command(label='Find', image=find_icon, compound=tk.LEFT, accelerator='Ctrl+F', command = find_func)


# help command
help.add_command(label='Help', image=help_icon, compound=tk.LEFT, command = about)

# redo and undo command
redo.add_command(label='Undo', image=undo_icon , compound=tk.LEFT, command = text_editor.edit_undo)
redo.add_command(label='Redo', image=redo_icon , compound=tk.LEFT, command = text_editor.edit_redo)

# view commands

show_statusbar = tk.BooleanVar()
show_statusbar.set(True)
show_toolbar = tk.BooleanVar()
show_toolbar.set(True)

def hide_toolbar():
    global show_toolbar
    if show_toolbar:
        tool_bar.pack_forget() # pack function will hide toolbar
        show_toolbar = False
    else :
        text_editor.pack_forget()
        status_bar.pack_forget()
        tool_bar.pack(side=tk.TOP, fill=tk.X)
        text_editor.pack(fill=tk.BOTH, expand=True)
        status_bar.pack(side=tk.BOTTOM)
        show_toolbar = True


def hide_statusbar():
    global show_statusbar
    if show_statusbar:
        status_bar.pack_forget()
        show_statusbar = False
    else :
        status_bar.pack(side=tk.BOTTOM)
        show_statusbar = True

view.add_checkbutton(label='Tool Bar',onvalue=True, offvalue=0,variable = show_toolbar, image=tool_bar_icon, compound=tk.LEFT, command=hide_toolbar)
view.add_checkbutton(label='Status Bar',onvalue=1, offvalue=False,variable = show_statusbar, image=status_bar_icon, compound=tk.LEFT, command=hide_statusbar)

# color theme loop
def change_theme():
    chosen_theme = theme_choice.get()
    color_tuple = color_dict.get(chosen_theme)
    fg_color, bg_color = color_tuple[0], color_tuple[1]
    text_editor.config(background=bg_color, fg=fg_color)

count = 0
for i in color_dict:
    color_theme.add_radiobutton(label=i, image=color_icons[count], variable=theme_choice, compound=tk.LEFT, command=change_theme)
    count +=1

######### main menu functionality end #######


main_application.config(menu=main_menu)

# binding shorcut keys
main_application.bind("<Control-n>", new_file)
main_application.bind("<Control-o>", open_file)
main_application.bind("<Control-s>", save_file)
main_application.bind("<Control-Alt-s>", save_as)
main_application.bind("<Control-q>", exit_func)
main_application.bind("<Control-f>", find_func)


main_application.mainloop()