import os
import extract_msg
import mimetypes
from email.message import EmailMessage
from tkinter import Tk, Label, Button, filedialog, messagebox

def msg_to_eml(msg_file_path, output_dir):
    try:
#import msg
        msg = extract_msg.Message(msg_file_path)
        msg_subject = msg.subject if msg.subject else "No Subject"
        msg_date = msg.date
        msg_sender = msg.sender
        msg_to = msg.to
        msg_body = msg.body
        msg_attachments = msg.attachments

#parse eml from import
        email = EmailMessage()
        email["Subject"] = msg_subject
        email["From"] = msg_sender
        email["To"] = msg_to
        email["Date"] = msg_date
        email.set_content(msg_body)

#attachment handling
        for attachment in msg_attachments:
            attachment_filename = attachment.longFilename or attachment.shortFilename or "unnamed"
            
#mimetype handling
            content_type, encoding = mimetypes.guess_type(attachment_filename)
            if content_type is None:
                content_type = 'application/octet-stream' 

            maintype, subtype = content_type.split('/', 1)

            email.add_attachment(
                attachment.data,
                maintype=maintype,
                subtype=subtype,
                filename=attachment_filename
            )

#export to eml
        eml_file_path = os.path.join(output_dir, f"{os.path.splitext(os.path.basename(msg_file_path))[0]}.eml")
        with open(eml_file_path, "wb") as eml_file:
            eml_file.write(email.as_bytes())

    except Exception as e:
        raise Exception(f"Error converting {msg_file_path}: {e}")

def batch_convert(input_directory, output_directory): #batch conversion and stuff
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    msg_files = [f for f in os.listdir(input_directory) if f.lower().endswith(".msg")]
    if not msg_files:
        messagebox.showinfo("No Files", "No .msg files found in the selected input directory.")
        return

    success_count = 0
    error_messages = []

    for file_name in msg_files:
        try:
            msg_to_eml(os.path.join(input_directory, file_name), output_directory)
            success_count += 1
        except Exception as e:
            error_messages.append(str(e))

    if success_count > 0:
        messagebox.showinfo("completed", f"{success_count} .msg files converted to .eml. just double check that they are working")

    if error_messages:
        error_text = "\n".join(error_messages)
        messagebox.showerror("Error", f"not all files converted compare whats in your selected output folder to the imput folder to see which ones didnt convert. set them aside and send this to andy:\n{error_text}")

#directory handling
def select_input_dir():
    directory = filedialog.askdirectory(title="select or make input folder")
    if directory:
        input_dir_label.config(text=directory)

def select_output_dir():
    directory = filedialog.askdirectory(title="select or make output folder")
    if directory:
        output_dir_label.config(text=directory)

#conversion
def start_conversion():
    input_dir = input_dir_label.cget("text")
    output_dir = output_dir_label.cget("text")

    if not input_dir or input_dir == "no input folder selected":
        messagebox.showerror("error", "select an input directory.")
        return

    if not output_dir or output_dir == "no output folder selected":
        messagebox.showerror("error", " select an output directory.")
        return

    batch_convert(input_dir, output_dir)

#interface
root = Tk()
root.title("uv water email converter")
root.resizable(False, False)

Label(root, text="input :").grid(row=0, column=0, padx=10, pady=10, sticky="w")
input_dir_label = Label(root, text="none selected", fg="gray", width=50, anchor="w")
input_dir_label.grid(row=0, column=1, padx=10, pady=10)
Button(root, text="browse", command=select_input_dir).grid(row=0, column=2, padx=10, pady=10)

Label(root, text="output :").grid(row=1, column=0, padx=10, pady=10, sticky="w")
output_dir_label = Label(root, text="none selected", fg="gray", width=50, anchor="w")
output_dir_label.grid(row=1, column=1, padx=10, pady=10)
Button(root, text="browse", command=select_output_dir).grid(row=1, column=2, padx=10, pady=10)

Button(root, text="start", command=start_conversion, bg="green", fg="white", width=15).grid(row=2, column=1, pady=20)

root.mainloop()