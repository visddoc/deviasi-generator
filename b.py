from tkinterdnd2 import TkinterDnD, DND_FILES
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import os
import threading
import sys
import subprocess

class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Merger Offline")
        self.root.geometry("900x600")
        self.root.resizable(True, True)

        self.pdf_files = []
        self.thumbnails = {}  # Dictionary untuk menyimpan thumbnail
        self.action_frame = None  # Frame untuk tombol action setelah merge

        self.create_widgets()
        self.setup_dnd()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Title
        ttk.Label(main_frame, text="PDF Merger - Offline", font=("Arial", 16, "bold")).pack(pady=10)

        # Input section
        input_frame = ttk.LabelFrame(main_frame, text="Add PDF Files")
        input_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(input_frame, text="Select PDF Files", command=self.add_files).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Label(input_frame, text="or Drag & Drop files here").pack(side=tk.LEFT, padx=5, pady=5)

        # File list & thumbnails
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.tree = ttk.Treeview(list_frame, columns=("Order", "File", "Pages"), show="headings", height=10)
        self.tree.heading("Order", text="Order")
        self.tree.heading("File", text="File Name")
        self.tree.heading("Pages", text="Pages")
        self.tree.column("Order", width=50, anchor=tk.CENTER)
        self.tree.column("File", width=400, anchor=tk.W)
        self.tree.column("Pages", width=80, anchor=tk.CENTER)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<<TreeviewSelect>>", self.show_preview)

        # Thumbnail preview
        preview_frame = ttk.LabelFrame(main_frame, text="Preview")
        preview_frame.pack(fill=tk.X, padx=5, pady=5)
        self.preview_label = ttk.Label(preview_frame, text="Select a file to preview first page")
        self.preview_label.pack(padx=5, pady=5)
        self.thumbnail_canvas = tk.Canvas(preview_frame, height=150, bg="#f0f0f0")
        self.thumbnail_canvas.pack(fill=tk.X, padx=5, pady=5)

        # Control buttons
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, padx=5, pady=10)
        ttk.Button(control_frame, text="Move Up", command=lambda: self.move_item(-1)).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Move Down", command=lambda: self.move_item(1)).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Remove Selected", command=self.remove_selected).pack(side=tk.LEFT, padx=5)

        # Output section
        output_frame = ttk.Frame(main_frame)
        output_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(output_frame, text="Output File:").pack(side=tk.LEFT, padx=5)
        self.output_var = tk.StringVar(value="merged.pdf")
        ttk.Entry(output_frame, textvariable=self.output_var, width=40).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(output_frame, text="Browse", command=self.browse_output).pack(side=tk.LEFT, padx=5)

        # Merge button
        ttk.Button(main_frame, text="Merge PDF", command=self.start_merge).pack(pady=10)

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W).pack(side=tk.BOTTOM, fill=tk.X)
        
        # Frame untuk tombol action setelah merge
        self.action_frame = ttk.Frame(main_frame)
        self.action_frame.pack(fill=tk.X, padx=10, pady=5)

    def setup_dnd(self):
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.handle_drop)

    def handle_drop(self, event):
        files = event.data.split() if sys.platform == 'win32' else self.root.tk.splitlist(event.data)
        pdf_files = [f for f in files if f.lower().endswith('.pdf')]
        self.add_pdf_files(pdf_files)

    def add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")], title="Select PDF files to merge")
        if files:
            self.add_pdf_files(files)

    def add_pdf_files(self, file_paths):
        for file_path in file_paths:
            if file_path not in self.pdf_files:
                try:
                    doc = fitz.open(file_path)
                    page_count = len(doc)
                    doc.close()
                    self.pdf_files.append(file_path)
                    self.tree.insert("", tk.END, values=(len(self.pdf_files), os.path.basename(file_path), page_count))
                    threading.Thread(target=self.generate_thumbnail, args=(file_path,), daemon=True).start()
                except Exception as e:
                    messagebox.showerror("Error", f"Could not open {file_path}:\n{str(e)}")
        self.status_var.set(f"Added {len(file_paths)} files. Total: {len(self.pdf_files)}")

    def generate_thumbnail(self, file_path):
        try:
            # Gunakan PyMuPDF untuk thumbnail
            doc = fitz.open(file_path)
            first_page = doc[0]  # Ambil halaman pertama
            
            # Buat thumbnail dengan ukuran yang sesuai
            pix = first_page.get_pixmap(matrix=fitz.Matrix(0.2, 0.2))
            
            # Konversi ke PIL Image
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img.thumbnail((120, 120))  # Resize ke ukuran thumbnail
            
            # Konversi ke PhotoImage untuk Tkinter
            photo = ImageTk.PhotoImage(img)
            self.thumbnails[file_path] = photo
            
            # Update preview jika file ini yang dipilih
            selected = self.tree.selection()
            if selected:
                selected_index = self.tree.index(selected[0])
                if selected_index < len(self.pdf_files) and self.pdf_files[selected_index] == file_path:
                    self.show_thumbnail(photo)
                    
            doc.close()
        except Exception as e:
            print(f"Error generating thumbnail for {file_path}: {str(e)}")

    def show_preview(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        try:
            index = self.tree.index(selected[0])
            if index < len(self.pdf_files):
                file_path = self.pdf_files[index]
                thumbnail = self.thumbnails.get(file_path)
                if thumbnail:
                    self.show_thumbnail(thumbnail)
                else:
                    self.preview_label.config(text="Generating preview...")
                    threading.Thread(target=self.generate_thumbnail, args=(file_path,), daemon=True).start()
        except Exception as e:
            print(f"Error showing preview: {str(e)}")

    def show_thumbnail(self, photo):
        self.thumbnail_canvas.delete("all")
        canvas_width = self.thumbnail_canvas.winfo_width()
        if canvas_width < 10:  # Jika canvas belum di-render
            canvas_width = 300  # Default width
        
        # Hitung posisi tengah
        x = (canvas_width - photo.width()) // 2
        
        self.thumbnail_canvas.create_image(x, 10, anchor=tk.NW, image=photo)
        self.thumbnail_canvas.image = photo
        self.preview_label.config(text="First page preview:")

    def move_item(self, direction):
        selected = self.tree.selection()
        if not selected:
            return
        try:
            index = self.tree.index(selected[0])
            new_index = index + direction
            if 0 <= new_index < len(self.pdf_files):
                # Tukar posisi file
                self.pdf_files[index], self.pdf_files[new_index] = self.pdf_files[new_index], self.pdf_files[index]
                
                # Refresh tampilan
                self.refresh_treeview()
                
                # Pilih item yang baru dipindahkan
                self.tree.selection_set(self.tree.get_children()[new_index])
        except Exception as e:
            print(f"Error moving item: {str(e)}")

    def remove_selected(self):
        selected = self.tree.selection()
        if not selected:
            return
        try:
            index = self.tree.index(selected[0])
            if index < len(self.pdf_files):
                removed_file = self.pdf_files.pop(index)
                if removed_file in self.thumbnails:
                    del self.thumbnails[removed_file]
                self.refresh_treeview()
                self.status_var.set(f"Removed 1 file. Total: {len(self.pdf_files)}")
        except Exception as e:
            print(f"Error removing item: {str(e)}")

    def refresh_treeview(self):
        # Hapus semua item di treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Tambahkan kembali semua file dengan urutan baru
        for i, file_path in enumerate(self.pdf_files):
            try:
                doc = fitz.open(file_path)
                page_count = len(doc)
                doc.close()
                self.tree.insert("", tk.END, values=(i+1, os.path.basename(file_path), page_count))
            except Exception as e:
                print(f"Error refreshing treeview for {file_path}: {str(e)}")
                self.tree.insert("", tk.END, values=(i+1, os.path.basename(file_path), "Error"))

    def browse_output(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", 
            filetypes=[("PDF files", "*.pdf")], 
            title="Save merged PDF as"
        )
        if file_path:
            self.output_var.set(file_path)

    def start_merge(self):
        if not self.pdf_files:
            messagebox.showwarning("Warning", "No PDF files to merge!")
            return
            
        output_path = self.output_var.get()
        if not output_path:
            messagebox.showwarning("Warning", "Please specify output file path!")
            return
            
        # Sembunyikan tombol action sebelumnya
        for widget in self.action_frame.winfo_children():
            widget.destroy()
            
        # Update UI
        self.status_var.set("Merging PDFs...")
        self.root.config(cursor="watch")
        self.root.update()
        
        # Mulai proses merge di thread terpisah
        threading.Thread(target=self.merge_pdfs, args=(output_path,), daemon=True).start()

    def merge_pdfs(self, output_path):
        try:
            merged_pdf = fitz.open()
            total_pages = 0
            
            # Merge semua file PDF
            for i, file_path in enumerate(self.pdf_files):
                pdf = fitz.open(file_path)
                merged_pdf.insert_pdf(pdf)
                total_pages += len(pdf)
                self.status_var.set(f"Merging file {i+1}/{len(self.pdf_files)}...")
                self.root.update()
                pdf.close()
                
            # Simpan hasil merge
            merged_pdf.save(output_path)
            merged_pdf.close()
            
            # Pesan sukses
            success_message = (
                f"PDF files merged successfully!\n\n"
                f"Total files: {len(self.pdf_files)}\n"
                f"Total pages: {total_pages}\n"
                f"Output file: {os.path.basename(output_path)}"
            )
            
            self.status_var.set(f"Success: {os.path.basename(output_path)}")
            messagebox.showinfo("Success", success_message)
            
            # Tampilkan tombol action
            self.show_success_actions(output_path)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to merge PDFs:\n{str(e)}")
            self.status_var.set("Merge failed")
        
        # Reset UI
        self.root.config(cursor="")
        self.root.update()

    def show_success_actions(self, output_path):
        # Hapus tombol sebelumnya
        for widget in self.action_frame.winfo_children():
            widget.destroy()
            
        # Tombol untuk membuka file
        open_btn = ttk.Button(
            self.action_frame, 
            text="Open PDF", 
            command=lambda: self.open_file(output_path)
        )
        open_btn.pack(side=tk.LEFT, padx=5)
        
        # Tombol untuk membuka folder
        folder_btn = ttk.Button(
            self.action_frame,
            text="Open Folder",
            command=lambda: self.open_folder(os.path.dirname(output_path))
        )
        folder_btn.pack(side=tk.LEFT, padx=5)
        
        # Tombol untuk merge lagi
        again_btn = ttk.Button(
            self.action_frame,
            text="Merge Again",
            command=self.reset_for_next_merge
        )
        again_btn.pack(side=tk.LEFT, padx=5)

    def open_file(self, file_path):
        try:
            if sys.platform == "win32":
                os.startfile(file_path)
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", file_path])
            else:  # Linux
                subprocess.run(["xdg-open", file_path])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file:\n{str(e)}")

    def open_folder(self, folder_path):
        try:
            if sys.platform == "win32":
                os.startfile(folder_path)
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", folder_path])
            else:  # Linux
                subprocess.run(["xdg-open", folder_path])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder:\n{str(e)}")

    def reset_for_next_merge(self):
        # Reset semua data
        self.pdf_files = []
        self.thumbnails = {}
        
        # Hapus semua item di treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Reset preview
        self.thumbnail_canvas.delete("all")
        self.preview_label.config(text="Select a file to preview first page")
        
        # Reset output filename
        self.output_var.set("merged.pdf")
        
        # Sembunyikan tombol action
        for widget in self.action_frame.winfo_children():
            widget.destroy()
            
        # Reset status
        self.status_var.set("Ready for next merge")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = PDFMergerApp(root)
    root.mainloop()