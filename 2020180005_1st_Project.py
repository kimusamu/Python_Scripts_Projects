import logging
logging.basicConfig(
level=logging.INFO,
format='%(filename)s(%(lineno)d): %(asctime)s - %(levelname)s - %(message)s'
)

from logging import info as info, debug as debug, warning as warning

from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
from tkinter import simpledialog
from PIL import Image, ImageTk
from rembg import remove

window = Tk()
window.title("My Project")
window.geometry('800x800')
window.resizable(True, True)

image = None
label = None
start_x = None
start_y = None
end_x = None
end_y = None

def open_image():
    global image, label

    file_name = filedialog.askopenfilename(
        title='Select Images',
        filetypes=(
            ("Png Images", '*.png'),
            ("Jpg Images", '*.jpg'),
            ("All Images", '*.*')
        )
    )

    if file_name:
        if label is not None:
            label.destroy()

        image = Image.open(file_name)
        photo = ImageTk.PhotoImage(image)
        label = Label(window, image=photo)
        label.image = photo
        label.pack()
        window.geometry(f'{image.size[0]}x{image.size[1]}')

def save_image():
    file_name = filedialog.asksaveasfilename(
        title='Save file as ...',
        filetypes=(
            ("Png Images", '*.png'),
            ("Jpg Images", '*.jpg'),
            ("All Images", '*.*')
        ),
        defaultextension='png'
    )

    if file_name:
        image.save(file_name)

def resize_image():
    global image, label

    new_width = simpledialog.askinteger("Input", "Enter new width:")
    new_height = simpledialog.askinteger("Input", "Enter new height:")

    image = image.resize((new_width, new_height))

    photo = ImageTk.PhotoImage(image)
    label.config(image=photo)
    label.image = photo
    window.geometry(f'{image.size[0]}x{image.size[1]}')

def delete_image_background():
    global image, label

    image = remove(image)
    photo = ImageTk.PhotoImage(image)
    label.config(image=photo)
    label.image = photo

def start_clip(event):
    global start_x, start_y
    start_x = event.x
    start_y = event.y

def end_clip(event):
    global image, label, start_x, start_y, end_x, end_y
    end_x = event.x
    end_y = event.y

    image = image.crop((start_x, start_y, end_x, end_y))
    photo = ImageTk.PhotoImage(image)
    label.config(image=photo)
    label.image = photo
    window.geometry(f'{image.size[0]}x{image.size[1]}')


def image_clip():
    window.bind("<Button-1>", start_clip)
    window.bind("<ButtonRelease-1>", end_clip)

def switch_left_right():
    global image, label

    image = image.transpose(Image.FLIP_LEFT_RIGHT)
    photo = ImageTk.PhotoImage(image)
    label.config(image=photo)
    label.image = photo

def convert_grayscale():
    global image, label

    image = image.convert("L")
    photo = ImageTk.PhotoImage(image)
    label.config(image=photo)
    label.image = photo

top_menu = Menu()

menu_File = Menu(master=top_menu)
menu_File.add_command(label='Open', accelerator='Ctrl+O', command=open_image)
menu_File.add_separator()
menu_File.add_command(label='Save', accelerator='Ctrl+S', command=save_image)
menu_File.add_separator()
menu_File.add_command(label='Quit', accelerator='Ctrl+Q', command=window.quit)
top_menu.add_cascade(label='File', menu=menu_File)
window.config(menu=top_menu)

menu_Tool = Menu(master=top_menu)
menu_Tool.add_command(label='Resize', accelerator='Ctrl+R', command=resize_image)
menu_Tool.add_separator()
menu_Tool.add_command(label='Clip', accelerator='Ctrl+C', command=image_clip)
menu_Tool.add_separator()
menu_Tool.add_command(label='Switch', accelerator='Ctrl+W', command=switch_left_right)
menu_Tool.add_separator()
menu_Tool.add_command(label='Grayscale', accelerator='Ctrl+G', command=convert_grayscale)
menu_Tool.add_separator()
menu_Tool.add_command(label='Delete BackGround', accelerator='Ctrl+D', command=delete_image_background)
top_menu.add_cascade(label='Tool', menu=menu_Tool)


# window.geometry("+%d+%d" % (window.winfo_screenwidth()/2, window.winfo_screenheight()/2))
# window.resizable(False, False)
window.mainloop()