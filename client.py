import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk as bootttk # Use alias for clarity
import socket
import os
import ttkbootstrap as ttk # Main theme provider
import zipfile
import tempfile
import traceback # For logging unexpected client errors

# Dependencies for Client: ttkbootstrap
# Dependencies for Server: Pillow, comtypes-client, pypdf, PyMuPDF,
#                          pdf2docx, python-pptx, reportlab
#                          (and MS Office if DOCX/PPTX/XLSX/HTML to PDF needed)

def client_program():
    HOST = '127.0.0.1'
    PORT = 65432

    # --- Helper Dialogs (Using Toplevel for modality and better widgets) ---
    def ask_password():
        dialog = tk.Toplevel(root)
        dialog.title("Password Required")
        dialog.geometry("300x120")
        dialog.transient(root); dialog.grab_set(); dialog.resizable(False, False)
        ttk.Label(dialog, text="Enter PDF password:").pack(pady=10)
        password_var = tk.StringVar()
        password_entry = ttk.Entry(dialog, textvariable=password_var, show='*')
        password_entry.pack(pady=5, padx=20, fill=tk.X)
        password_entry.focus_set()
        result = {"password": None}
        def on_ok(): result["password"] = password_var.get(); dialog.destroy()
        def on_cancel(): dialog.destroy()
        button_frame = ttk.Frame(dialog)
        ttk.Button(button_frame, text="OK", command=on_ok, bootstyle="primary").pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Cancel", command=on_cancel).pack(side=tk.RIGHT, padx=10)
        button_frame.pack(pady=10)
        # Bind Enter key to OK
        dialog.bind('<Return>', lambda event=None: on_ok())
        root.wait_window(dialog)
        return result["password"] # None if cancelled

    def ask_split_ranges():
        # Use simpledialog for this one as it's just a string
        ranges = simpledialog.askstring("Split PDF", "Enter page ranges (e.g., 1-3, 5, 7-):", parent=root)
        return ranges

    def ask_rotate_options():
        dialog = tk.Toplevel(root)
        dialog.title("Rotate PDF Options")
        dialog.geometry("300x180"); dialog.transient(root); dialog.grab_set(); dialog.resizable(False, False)
        result = {"pages": None, "angle": None}
        # Pages Entry
        ttk.Label(dialog, text="Pages to rotate (e.g., 1, 3-5, all):").pack(pady=5, padx=10, anchor='w')
        pages_var = tk.StringVar(value="all") # Default to all
        pages_entry = ttk.Entry(dialog, textvariable=pages_var)
        pages_entry.pack(pady=2, padx=10, fill=tk.X)
        pages_entry.focus_set()
        # Angle Selection
        ttk.Label(dialog, text="Rotation angle:").pack(pady=5, padx=10, anchor='w')
        angle_var = tk.IntVar(value=90)
        angle_combo = bootttk.Combobox(dialog, textvariable=angle_var, values=[90, 180, 270, -90], state="readonly", width=10)
        angle_combo.pack(pady=2, padx=10)
        def on_ok(): result["pages"] = pages_var.get(); result["angle"] = angle_var.get(); dialog.destroy()
        def on_cancel(): dialog.destroy()
        button_frame = ttk.Frame(dialog); button_frame.pack(pady=15)
        ttk.Button(button_frame, text="OK", command=on_ok, bootstyle="primary").pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Cancel", command=on_cancel).pack(side=tk.RIGHT, padx=10)
        dialog.bind('<Return>', lambda event=None: on_ok())
        root.wait_window(dialog)
        return result["pages"], result["angle"] # Returns None, None if cancelled

    def ask_page_number_position():
         dialog = tk.Toplevel(root); dialog.title("Page Number Position"); dialog.geometry("300x130")
         dialog.transient(root); dialog.grab_set(); dialog.resizable(False, False)
         result = {"position": None}
         ttk.Label(dialog, text="Select position:").pack(pady=5)
         positions = ['bottom-center', 'bottom-left', 'bottom-right', 'top-center', 'top-left', 'top-right']
         position_var = tk.StringVar(value=positions[0])
         position_combo = bootttk.Combobox(dialog, textvariable=position_var, values=positions, state="readonly", width=18)
         position_combo.pack(pady=5)
         position_combo.focus_set()
         def on_ok(): result["position"] = position_var.get(); dialog.destroy()
         def on_cancel(): dialog.destroy()
         button_frame = ttk.Frame(dialog); button_frame.pack(pady=15)
         ttk.Button(button_frame, text="OK", command=on_ok, bootstyle="primary").pack(side=tk.LEFT, padx=10)
         ttk.Button(button_frame, text="Cancel", command=on_cancel).pack(side=tk.RIGHT, padx=10)
         dialog.bind('<Return>', lambda event=None: on_ok())
         root.wait_window(dialog)
         return result["position"] # None if cancelled

    # --- Main Communication Logic ---
    def send_request_to_server(action, file_path_or_paths, options=None):
        is_merge = action == "merge"
        local_zip_to_send = None
        sock = None
        options = options or {} # Ensure options is a dict

        try:
            # Establish Connection
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(30.0) # Shorter timeout for initial connect
            print(f"Connecting to {HOST}:{PORT}...")
            sock.connect((HOST, PORT))
            print(f"Connected to server for action: {action}")
            sock.settimeout(600.0) # Longer timeout for operations/transfer

            # 1. Send Action
            print(f"Sending action: {action}")
            sock.sendall(action.encode())
            ack = sock.recv(1024); print(f"Recv ACK: {ack}")
            if ack != b'ACK_ACTION': raise ConnectionAbortedError("Invalid ACK after sending action.")

            # --- Prepare File Info ---
            if is_merge:
                if not isinstance(file_path_or_paths, (list, tuple)) or len(file_path_or_paths) < 2:
                    raise ValueError("Merge action requires at least two files.")
                with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as temp_zip:
                    local_zip_to_send = temp_zip.name
                    print(f"Creating temporary zip for merge: {local_zip_to_send}")
                    valid_files_count = 0
                    with zipfile.ZipFile(temp_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        for file_path in file_path_or_paths:
                             if os.path.isfile(file_path): zipf.write(file_path, os.path.basename(file_path)); valid_files_count+=1
                             else: print(f"Warning: Skipping non-existent file for merge: {file_path}")
                    if valid_files_count < 2 : raise ValueError(f"Merge requires at least two valid files (found {valid_files_count}).")
                file_to_send = local_zip_to_send
                filename_for_server = f"merge_input_{valid_files_count}files.zip"
            else:
                if not isinstance(file_path_or_paths, str) or not os.path.isfile(file_path_or_paths):
                     raise ValueError(f"Input file not found or invalid: {file_path_or_paths}")
                file_to_send = file_path_or_paths
                filename_for_server = os.path.basename(file_to_send)
            # --- End Prepare File Info ---

            # 2. Send Base Filename
            print(f"Sending base filename: {filename_for_server}")
            sock.sendall(filename_for_server.encode())
            ack = sock.recv(1024); print(f"Recv ACK: {ack}")
            if ack != b'ACK_FILENAME': raise ConnectionAbortedError("Invalid ACK after sending filename.")

            # 3. Send File Size & Data
            file_size = os.path.getsize(file_to_send)
            print(f"Sending file size: {file_size}")
            sock.sendall(str(file_size).encode().ljust(16))
            ack = sock.recv(1024); print(f"Recv ACK: {ack}")
            if ack != b'ACK_SIZE': raise ConnectionAbortedError("Invalid ACK after sending size.")
            print(f"Sending file data from: {file_to_send}...")
            with open(file_to_send, "rb") as f:
                 bytes_sent = 0
                 while True:
                     chunk = f.read(65536);
                     if not chunk: break
                     sock.sendall(chunk); bytes_sent += len(chunk)
            print(f"Sent {bytes_sent} bytes of file data.")

            # 4. Send Extra Options (if required by action)
            if action in ("encrypt", "decrypt"):
                password = options.get("password") # Get from options dict
                if password is None: raise ValueError("Password not provided for encrypt/decrypt.")
                print("Sending password...")
                sock.sendall(password.encode()); ack = sock.recv(1024); print(f"Recv ACK: {ack}")
                if ack != b'ACK_PASS': raise ConnectionAbortedError("Invalid ACK after sending password.")
            elif action == "split":
                ranges_str = options.get("ranges")
                if not ranges_str: raise ValueError("Split ranges not provided.")
                print(f"Sending split ranges: {ranges_str}")
                sock.sendall(ranges_str.encode()); ack = sock.recv(1024); print(f"Recv ACK: {ack}")
                if ack != b'ACK_RANGES': raise ConnectionAbortedError("Invalid ACK after sending ranges.")
            elif action == "rotate":
                pages_str = options.get("pages"); angle_str = str(options.get("angle"))
                if not pages_str or not angle_str: raise ValueError("Rotation pages/angle not provided.")
                print(f"Sending rotate pages: {pages_str}")
                sock.sendall(pages_str.encode()); ack = sock.recv(1024); print(f"Recv ACK: {ack}")
                if ack != b'ACK_PAGES': raise ConnectionAbortedError("Invalid ACK after sending pages.")
                print(f"Sending rotate angle: {angle_str}")
                sock.sendall(angle_str.encode()); ack = sock.recv(1024); print(f"Recv ACK: {ack}")
                if ack != b'ACK_ANGLE': raise ConnectionAbortedError("Invalid ACK after sending angle.")
            elif action == "add_numbers":
                position = options.get("position")
                if not position: raise ValueError("Page number position not provided.")
                print(f"Sending page number position: {position}")
                sock.sendall(position.encode()); ack = sock.recv(1024); print(f"Recv ACK: {ack}")
                if ack != b'ACK_POSITION': raise ConnectionAbortedError("Invalid ACK after sending position.")

            # --- Receive Result ---
            print("Waiting for result from server...")

            # 5. Receive Output Filename Suggestion
            output_filename_suggestion_bytes = sock.recv(1024)
            if not output_filename_suggestion_bytes: raise ConnectionAbortedError("Server disconnected before sending output filename.")
            output_filename_suggestion = output_filename_suggestion_bytes.decode()
            sock.sendall(b"ACK_OUT_FILENAME"); print("Sent ACK_OUT_FILENAME")
            print(f"Received suggested output filename: {output_filename_suggestion}")

            # 6. Receive Output File Size
            output_size_bytes = sock.recv(16)
            if not output_size_bytes: raise ConnectionAbortedError("Server disconnected before sending output size.")
            output_size = int(output_size_bytes.decode().strip())
            sock.sendall(b"ACK_OUT_SIZE"); print("Sent ACK_OUT_SIZE")
            print(f"Expecting output size: {output_size} bytes")

            # Check for server error signal (0 size + error name)
            if output_size == 0 and "error_" in output_filename_suggestion.lower():
                error_msg_bytes = sock.recv(4096) # Receive potential error message
                error_msg = error_msg_bytes.decode(errors='ignore')
                raise Exception(f"Server Error: {error_msg}" if error_msg else f"Server indicated an error ({output_filename_suggestion}).")
            elif output_size == 0:
                # Valid empty file (e.g., split resulted in nothing for ranges)
                 messagebox.showwarning("Empty Result", f"Server returned an empty result file for action '{action}'.\nFilename: {output_filename_suggestion}", parent=root)
                 # No need to receive data or save. Indicate completion.
                 return True # Indicate success but with empty result

            # 7. Receive Output File Data
            print(f"Receiving result data ({output_size} bytes)...")
            output_data = b""
            received_bytes = 0
            while received_bytes < output_size:
                chunk = sock.recv(min(65536, output_size - received_bytes))
                if not chunk: raise ConnectionAbortedError(f"Server disconnected during result transfer ({received_bytes}/{output_size} received).")
                output_data += chunk; received_bytes += len(chunk)
            print(f"Received {len(output_data)} bytes of result data.")

            # --- Save Result ---
            default_ext = os.path.splitext(output_filename_suggestion)[1] or ".bin" # Ensure default ext
            file_types = []
            type_map = {".pdf": "PDF files", ".zip": "ZIP archives", ".docx": "Word Documents", ".pptx": "PowerPoint Presentations"}
            desc = type_map.get(default_ext.lower())
            if desc: file_types.append((desc, f"*{default_ext.lower()}"))
            file_types.append(("All files", "*.*"))

            save_path = filedialog.asksaveasfilename(parent=root, title=f"Save Result ({action})",
                initialfile=output_filename_suggestion, defaultextension=default_ext, filetypes=file_types)

            if save_path:
                with open(save_path, "wb") as f: f.write(output_data)
                messagebox.showinfo("Success", f"File processed and saved successfully!\nPath: {save_path}", parent=root)
                return True # Indicate success
            else:
                messagebox.showwarning("Cancelled", "Save operation cancelled by user.", parent=root)
                return False # Indicate cancellation

        except socket.timeout: messagebox.showerror("Timeout Error", "Connection or operation timed out.", parent=root); return False
        except ConnectionRefusedError: messagebox.showerror("Connection Error", "Could not connect to the server.", parent=root); return False
        except (ConnectionAbortedError, ConnectionResetError) as cae: messagebox.showerror("Connection Error", f"Connection lost:\n{cae}", parent=root); return False
        except ValueError as ve: messagebox.showerror("Input Error", f"{ve}", parent=root); return False
        except Exception as e:
            print(f"An unexpected client error occurred:\n{traceback.format_exc()}")
            messagebox.showerror("Error", f"An unexpected error occurred:\n{e}", parent=root); return False
        finally:
            if local_zip_to_send and os.path.exists(local_zip_to_send):
                 try: os.remove(local_zip_to_send); print(f"Removed temporary zip: {local_zip_to_send}")
                 except OSError as e_rem: print(f"Warning: Could not remove temp zip {local_zip_to_send}: {e_rem}")
            if sock:
                try: sock.shutdown(socket.SHUT_RDWR)
                except OSError: pass
                finally: sock.close(); print("Socket closed.")
        return False # Indicate failure if exception occurred

    # --- UI Button Callbacks ---
    def update_label_status(label, text):
         """Helper to update label and force UI refresh."""
         if isinstance(label, ttk.Label): # Check if it's a valid label widget
            label.config(text=text)
            root.update_idletasks()
         else:
             print(f"Warning: Attempted to update non-label widget: {label}")

    # Generic file upload handler, now focuses on getting path and calling server comms
    def handle_upload(file_types, label, action, options_func=None, is_multiple=False, title_suffix=""):
        original_text = label.cget("text")
        file_path_or_paths = None
        opts = {}

        try: # Add try block for file dialogs and options gathering
            if is_multiple:
                file_path_or_paths = filedialog.askopenfilenames(parent=root, title=f"Select Files for {title_suffix}", filetypes=file_types)
                if not file_path_or_paths or len(file_path_or_paths) < 2:
                    messagebox.showwarning("Selection Error", "Please select at least two files for merging.", parent=root)
                    label.config(text=original_text); return # Restore label and exit
            else:
                file_path_or_paths = filedialog.askopenfilename(parent=root, title=f"Select File for {title_suffix}", filetypes=file_types)
                if not file_path_or_paths: label.config(text=original_text); return # Restore label if cancelled

            # Gather options if an options function is provided
            if options_func:
                opts = options_func()
                if opts is None or any(v is None for v in opts.values()): # Check if options were cancelled
                    messagebox.showinfo("Cancelled", f"{title_suffix} operation cancelled.", parent=root)
                    label.config(text=original_text); return # Restore label

            # Update label before potentially long operation
            display_name = f"{len(file_path_or_paths)} files" if is_multiple else os.path.basename(file_path_or_paths)
            update_label_status(label, f"Processing: {display_name}...")

            # Call the server communication function
            success = send_request_to_server(action, file_path_or_paths, options=opts)

            # Update label based on success/failure/cancel
            if success is True:
                update_label_status(label, f"Completed: {display_name}")
            elif success is False and not is_multiple: # Check if save was cancelled (False returned)
                 update_label_status(label, f"Selected: {display_name} (Save Cancelled)")
            elif success is False and is_multiple:
                 update_label_status(label, f"Selected: {display_name} (Merge Cancelled/Failed)")
            else: # Handle unexpected errors from send_request
                update_label_status(label, f"Failed: {display_name}")

        except Exception as e:
             print(f"Error during handle_upload for {action}: {traceback.format_exc()}")
             messagebox.showerror("Client Error", f"Failed during UI operation:\n{e}", parent=root)
             label.config(text=original_text) # Restore label on error


    # Simplified option gathering functions that return dicts or None
    def get_encrypt_decrypt_opts(): pwd = ask_password(); return {"password": pwd} if pwd is not None else None
    def get_split_opts(): ranges = ask_split_ranges(); return {"ranges": ranges} if ranges is not None else None
    def get_rotate_opts(): pages, angle = ask_rotate_options(); return {"pages": pages, "angle": angle} if pages is not None else None
    def get_add_numbers_opts(): pos = ask_page_number_position(); return {"position": pos} if pos is not None else None


    # --- UI Setup ---
    root = ttk.Window(title="Advanced File Converter & PDF Tools v1.1", themename="darkly")
    root.geometry("750x750")
    root.columnconfigure(0, weight=1); root.rowconfigure(0, weight=1)

    base_frame = ttk.Frame(root); base_frame.grid(row=0, column=0, sticky="nsew")
    base_frame.rowconfigure(0, weight=1); base_frame.columnconfigure(0, weight=1)
    canvas = tk.Canvas(base_frame, borderwidth=0, highlightthickness=0)
    scrollbar = ttk.Scrollbar(base_frame, orient="vertical", command=canvas.yview)
    main_frame = ttk.Frame(canvas, padding=20) # Frame content goes here
    main_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=main_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.grid(row=0, column=0, sticky="nsew"); scrollbar.grid(row=0, column=1, sticky="ns")
    main_frame.grid_columnconfigure(0, weight=1); main_frame.grid_columnconfigure(1, weight=0)

    # --- UI Elements Creation ---
    label_style = {"anchor": "w", "wraplength": 450}; button_style = "info"
    row_num = 0

    def add_option_row(parent, display_text, command_func, initial_label_text=None, button_bootstyle=button_style):
         nonlocal row_num
         if initial_label_text is None: initial_label_text = f"No file selected for {display_text}"
         label = ttk.Label(parent, text=initial_label_text, width=60, **label_style)
         label.grid(row=row_num, column=0, sticky="ew", padx=(0, 10), pady=4)
         button = ttk.Button(parent, text=display_text, width=20, command=lambda l=label: command_func(l), bootstyle=button_bootstyle)
         button.grid(row=row_num, column=1, sticky="e", pady=4)
         row_num += 1

    def add_section_header(parent, text):
        nonlocal row_num
        ttk.Separator(parent, orient='horizontal').grid(row=row_num, column=0, columnspan=2, sticky='ew', pady=(15, 5)); row_num += 1
        ttk.Label(parent, text=text, font=("Arial", 12, "bold")).grid(row=row_num, column=0, columnspan=2, sticky='w', pady=(0, 10)); row_num += 1

    # --- Define UI Sections ---
    add_section_header(main_frame, "File to PDF Conversions")
    add_option_row(main_frame, "JPG/PNG to PDF", lambda l: handle_upload([("Image files", "*.jpg *.jpeg *.png")], l, "convert", title_suffix="JPG/PNG to PDF"))
    add_option_row(main_frame, "DOCX to PDF", lambda l: handle_upload([("Word files", "*.docx")], l, "convert", title_suffix="DOCX to PDF"))
    add_option_row(main_frame, "PPTX to PDF", lambda l: handle_upload([("PowerPoint files", "*.pptx")], l, "convert", title_suffix="PPTX to PDF"))
    add_option_row(main_frame, "XLSX to PDF", lambda l: handle_upload([("Excel files", "*.xlsx")], l, "convert", title_suffix="XLSX to PDF"))
    add_option_row(main_frame, "HTML to PDF", lambda l: handle_upload([("HTML files", "*.html *.htm")], l, "convert", title_suffix="HTML to PDF"))

    add_section_header(main_frame, "PDF Security")
    add_option_row(main_frame, "Encrypt PDF", lambda l: handle_upload([("PDF files", "*.pdf")], l, "encrypt", options_func=get_encrypt_decrypt_opts, title_suffix="Encrypt PDF"), initial_label_text="No PDF selected for Encryption")
    add_option_row(main_frame, "Decrypt PDF", lambda l: handle_upload([("PDF files", "*.pdf")], l, "decrypt", options_func=get_encrypt_decrypt_opts, title_suffix="Decrypt PDF"), initial_label_text="No PDF selected for Decryption")

    add_section_header(main_frame, "PDF to Other Formats")
    pdf_types = [("PDF files", "*.pdf")]
    add_option_row(main_frame, "PDF to JPG (ZIP)", lambda l: handle_upload(pdf_types, l, "pdf_to_jpg", title_suffix="PDF to JPG"), initial_label_text="No PDF selected for PDF to JPG")
    add_option_row(main_frame, "PDF to WORD (.docx)", lambda l: handle_upload(pdf_types, l, "pdf_to_word", title_suffix="PDF to Word"), initial_label_text="No PDF selected for PDF to Word")
    add_option_row(main_frame, "PDF to PPTX (Images)", lambda l: handle_upload(pdf_types, l, "pdf_to_pptx", title_suffix="PDF to PPTX"), initial_label_text="No PDF selected for PDF to PPTX")

    add_section_header(main_frame, "PDF Manipulation Tools")
    add_option_row(main_frame, "Compress PDF", lambda l: handle_upload(pdf_types, l, "compress", title_suffix="Compress PDF"), initial_label_text="No PDF selected for Compress PDF")
    add_option_row(main_frame, "Split PDF", lambda l: handle_upload(pdf_types, l, "split", options_func=get_split_opts, title_suffix="Split PDF"), initial_label_text="No PDF selected for Split PDF")
    add_option_row(main_frame, "Merge PDFs", lambda l: handle_upload(pdf_types, l, "merge", is_multiple=True, title_suffix="Merge PDFs"), initial_label_text="No files selected for Merge PDF")
    add_option_row(main_frame, "Rotate PDF", lambda l: handle_upload(pdf_types, l, "rotate", options_func=get_rotate_opts, title_suffix="Rotate PDF"), initial_label_text="No PDF selected for Rotate PDF")
    add_option_row(main_frame, "Add Page Numbers", lambda l: handle_upload(pdf_types, l, "add_numbers", options_func=get_add_numbers_opts, title_suffix="Add Page Numbers"), initial_label_text="No PDF selected for Add Page Numbers")

    root.mainloop()

if __name__ == '__main__':
    client_program()