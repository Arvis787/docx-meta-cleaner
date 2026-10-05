from __future__ import annotations

import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from cleaner import OFFICE_EXTENSIONS, is_office_file, iter_office_files, remove_office_metadata
from settings import APP_NAME, desktop_directory, dialog_directory, load_settings, restored_selection, save_settings, settings_file_path

APP_VERSION = '1.2.0'
FORMATS = '*.docx *.docm *.dotx *.dotm *.pptx *.pptm *.potx *.potm *.ppsx *.ppsm *.xlsx *.xlsm *.xltx *.xltm'


class DocxMetaCleanerApp(tk.Tk):
    def __init__(self, settings_path: Path | None = None):
        super().__init__()
        self.title(f'{APP_NAME} {APP_VERSION}')
        self.geometry('860x640')
        self.minsize(760, 560)
        self.configure(bg='#f4f7fb')
        self.settings_path = settings_path or settings_file_path()
        self.settings = load_settings(self.settings_path)
        self.selected_path, mode = restored_selection(self.settings)
        self.selection_mode = tk.StringVar(value=mode)
        self.path_var = tk.StringVar(value=str(self.selected_path) if self.selected_path else 'Файл или папка не выбраны')
        self.backup_var = tk.BooleanVar(value=self.settings.get('make_backup', True) is not False)
        self.recursive_var = tk.BooleanVar(value=self.settings.get('recursive', True) is not False)
        self.status_var = tk.StringVar(value='Готово к работе')
        self.processing = False
        self.messages = queue.Queue()
        self.controls = []
        self.poll_id = None
        asset_dir = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent)) / 'assets'
        try:
            self.app_icon = tk.PhotoImage(file=str(asset_dir / 'app.png'))
            self.iconphoto(True, self.app_icon)
        except tk.TclError:
            pass
        self._build_ui()
        self.protocol('WM_DELETE_WINDOW', self.close)
        self.poll_id = self.after(100, self._poll_messages)

    def _build_ui(self):
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('TFrame', background='#f4f7fb')
        style.configure('TLabel', background='#f4f7fb', foreground='#172033', font=('Segoe UI', 10))
        style.configure('TButton', font=('Segoe UI', 10), padding=8)
        style.configure('TCheckbutton', background='#f4f7fb', font=('Segoe UI', 10))
        header = tk.Frame(self, bg='#2563eb')
        header.pack(fill='x')
        tk.Label(header, text=APP_NAME, bg='#2563eb', fg='white', font=('Segoe UI', 20, 'bold')).pack(anchor='w', padx=24, pady=(18, 3))
        tk.Label(header, text=f'Word  /  PowerPoint  /  Excel                       v{APP_VERSION}', bg='#2563eb', fg='white', font=('Segoe UI', 10)).pack(anchor='w', padx=24, pady=(0, 18))
        body = ttk.Frame(self, padding=24)
        body.pack(fill='both', expand=True)
        toolbar = ttk.Frame(body)
        toolbar.pack(fill='x', pady=(0, 16))
        for label, command in [('Выбрать файл', self.select_file), ('Выбрать папку', self.select_folder), ('Рабочий стол', self.select_desktop)]:
            button = ttk.Button(toolbar, text=label, command=command)
            button.pack(side='left', padx=(0, 12))
            self.controls.append(button)
        ttk.Label(body, text='Выбрано', font=('Segoe UI', 10, 'bold')).pack(anchor='w')
        path_label = ttk.Label(body, textvariable=self.path_var, wraplength=790)
        path_label.pack(fill='x', pady=(6, 14))
        body.bind('<Configure>', lambda event: path_label.configure(wraplength=max(100, event.width - 48)))
        for label, variable in [('Создавать резервные копии .bak', self.backup_var), ('Обрабатывать подпапки', self.recursive_var)]:
            control = ttk.Checkbutton(body, text=label, variable=variable, command=self._save_settings)
            control.pack(anchor='w', pady=3)
            self.controls.append(control)
        self.clean_button = tk.Button(body, text='Очистить метаданные', command=self.start_cleaning,
                                      bg='#2563eb', activebackground='#1d4ed8', fg='white', activeforeground='white',
                                      font=('Segoe UI', 11, 'bold'), relief='flat', pady=12)
        self.clean_button.pack(fill='x', pady=(16, 12))
        self.controls.append(self.clean_button)
        ttk.Label(body, textvariable=self.status_var).pack(anchor='w', pady=(0, 6))
        self.progress = ttk.Progressbar(body, mode='determinate')
        self.progress.pack(fill='x', pady=(0, 16))
        log_header = ttk.Frame(body)
        log_header.pack(fill='x', pady=(0, 6))
        ttk.Label(log_header, text='Журнал', font=('Segoe UI', 10, 'bold')).pack(side='left')
        ttk.Button(log_header, text='Очистить журнал', command=self.clear_log).pack(side='right')
        log_frame = ttk.Frame(body)
        log_frame.pack(fill='both', expand=True)
        self.log_text = tk.Text(log_frame, height=8, bg='#172033', fg='#e5e7eb', font=('Consolas', 10), wrap='word', relief='flat', padx=12, pady=10, state='disabled')
        self.log_text.pack(side='left', fill='both', expand=True)
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side='right', fill='y')
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.tag_configure('error', foreground='#fda4af')
        self.log_text.tag_configure('ok', foreground='#86efac')

    def _save_settings(self):
        data = dict(self.settings)
        data.update(recursive=self.recursive_var.get(), make_backup=self.backup_var.get())
        if self.selected_path:
            data.update(selected_path=str(self.selected_path), selection_mode=self.selection_mode.get(),
                        last_directory=str(self.selected_path if self.selected_path.is_dir() else self.selected_path.parent))
        self.settings = data
        try:
            save_settings(self.settings_path, data)
        except OSError as error:
            self.log(f'Не удалось сохранить настройки: {error}', 'error')

    def _select(self, selected: str | Path, mode: str):
        path = Path(selected).expanduser().resolve()
        if mode == 'folder' and not path.is_dir():
            messagebox.showerror(APP_NAME, 'Папка не существует.', parent=self)
            return
        if mode == 'file' and not is_office_file(path):
            messagebox.showerror(APP_NAME, 'Выберите поддерживаемый файл Word, PowerPoint или Excel.', parent=self)
            return
        self.selected_path = path
        self.selection_mode.set(mode)
        self.path_var.set(str(path))
        self.status_var.set('Выбрана папка' if mode == 'folder' else 'Выбран файл')
        self._save_settings()

    def select_file(self):
        if self.processing:
            return
        selected = filedialog.askopenfilename(parent=self, title='Выберите файл Office', initialdir=dialog_directory(self.settings), filetypes=[
            ('Word, PowerPoint и Excel', FORMATS), ('Word', '*.docx *.docm *.dotx *.dotm'),
            ('PowerPoint', '*.pptx *.pptm *.potx *.potm *.ppsx *.ppsm'), ('Excel', '*.xlsx *.xlsm *.xltx *.xltm'),
        ])
        if selected:
            self._select(selected, 'file')

    def select_folder(self):
        if self.processing:
            return
        selected = filedialog.askdirectory(parent=self, title='Выберите папку с файлами Office', initialdir=dialog_directory(self.settings), mustexist=True)
        if selected:
            self._select(selected, 'folder')

    def select_desktop(self):
        if not self.processing:
            self._select(desktop_directory(), 'folder')

    def start_cleaning(self):
        if self.processing:
            return
        if not self.selected_path:
            messagebox.showinfo(APP_NAME, 'Выберите файл или папку.', parent=self)
            return
        if self.selection_mode.get() == 'file' and not is_office_file(self.selected_path):
            messagebox.showerror(APP_NAME, 'Выбранный файл больше не доступен.', parent=self)
            return
        if self.selection_mode.get() == 'folder' and not self.selected_path.is_dir():
            messagebox.showerror(APP_NAME, 'Выбранная папка больше не доступна.', parent=self)
            return
        self._save_settings()
        self.processing = True
        for control in self.controls:
            control.configure(state='disabled')
        self.status_var.set('Поиск файлов...')
        self.progress.configure(value=0)
        threading.Thread(target=self._clean_worker, args=(self.selected_path, self.selection_mode.get(), self.recursive_var.get(), self.backup_var.get()), daemon=True).start()

    def _clean_worker(self, path: Path, mode: str, recursive: bool, backup: bool):
        success = errors = 0
        try:
            files = [path] if mode == 'file' else sorted(iter_office_files(path, recursive))
            self.messages.put(('total', len(files)))
            for index, file in enumerate(files, 1):
                try:
                    result = remove_office_metadata(file, make_backup=backup)
                    suffix = f'; копия: {result.backup_path.name}' if result.backup_path else ''
                    self.messages.put(('log', f'{file}: удалено записей {result.removed_items}{suffix}', 'ok'))
                    success += 1
                except Exception as error:
                    self.messages.put(('log', f'{file}: {error}', 'error'))
                    errors += 1
                self.messages.put(('progress', index))
        except Exception as error:
            self.messages.put(('log', str(error), 'error'))
            errors += 1
        finally:
            self.messages.put(('done', success, errors))

    def _poll_messages(self):
        try:
            while True:
                item = self.messages.get_nowait()
                if item[0] == 'log':
                    self.log(item[1], item[2])
                elif item[0] == 'total':
                    self.progress.configure(maximum=max(1, item[1]), value=0)
                    self.status_var.set(f'Найдено файлов: {item[1]}')
                elif item[0] == 'progress':
                    self.progress.configure(value=item[1])
                elif item[0] == 'done':
                    self.processing = False
                    for control in self.controls:
                        control.configure(state='normal')
                    status = f'Обработано: {item[1]}; ошибок: {item[2]}' if item[1] or item[2] else 'Поддерживаемые файлы не найдены'
                    self.status_var.set(status)
                    self.log(status)
        except queue.Empty:
            pass
        self.poll_id = self.after(100, self._poll_messages)

    def log(self, message: str, tag: str = ''):
        self.log_text.configure(state='normal')
        self.log_text.insert('end', message + '\n', tag)
        self.log_text.see('end')
        self.log_text.configure(state='disabled')

    def clear_log(self):
        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', 'end')
        self.log_text.configure(state='disabled')

    def close(self):
        if self.processing:
            messagebox.showinfo(APP_NAME, 'Дождитесь окончания обработки файлов.', parent=self)
            return
        self._save_settings()
        if self.poll_id:
            self.after_cancel(self.poll_id)
        self.destroy()


if __name__ == '__main__':
    DocxMetaCleanerApp().mainloop()
