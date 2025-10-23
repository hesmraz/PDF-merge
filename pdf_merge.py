import sys, os, time
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QFileDialog,
    QLabel, QVBoxLayout, QMessageBox, QSlider, QHBoxLayout, QScrollArea, QSplitter
)
from PyQt5.QtGui import QPixmap, QPainter, QPen
from PyQt5.QtCore import Qt, QRect
import fitz  # PyMuPDF
from pdf2image import convert_from_path
from PIL import Image

class PDFMergerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Merger")
        self.setMinimumWidth(900)
        self.setMinimumHeight(700)
        
        # Создаём папку temp если её нет
        self.temp_dir = "temp"
        if not os.path.exists(self.temp_dir):
            os.makedirs(self.temp_dir)
        
        self.template_path = ""
        self.insert_pdf_path = ""
        self.preview_path = ""
        self.insert_rect = None
        self.insert_area_chosen = False

        self.start_point = None
        self.end_point = None
        self.template_pixmap = None

        self.crop_path = os.path.join(self.temp_dir, "crop_preview.png")
        self.sample_path = ""
        self.crop_rect = None
        self.crop_chosen = False

        self.insert_pos = None   # x, y pos on preview
        self.insert_width = 150  # Default insert width in px (on preview)
        self.insert_height = 150 # Default insert height in px (on preview)
        self.insert_aspect_ratio = 1.0  # width/height
        self.dragging = False
        self.slider = None   # QSlider for insert size

        self.insert_images_ready = False
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        self.template_btn = QPushButton("Выбрать PDF шаблон")
        self.insert_btn = QPushButton("Выбрать PDF вставку")
        self.merge_btn = QPushButton("Слияние")
        self.status_label = QLabel("1. Выбери PDF шаблон. 2. PDF вставку. 3. Обрежь область на вставке. 4. Слияние.")

        self.template_btn.clicked.connect(self.choose_template)
        self.insert_btn.clicked.connect(self.choose_insert_pdf)
        self.merge_btn.clicked.connect(self.start_merge)

        layout.addWidget(self.template_btn)
        layout.addWidget(self.insert_btn)
        layout.addWidget(self.status_label)

        # Создаём QSplitter вертикальный
        splitter = QSplitter(Qt.Vertical)

        # Crop preview для вставки в ScrollArea
        self.crop_label = QLabel("Появится вставка для обрезки")
        self.crop_label.setAlignment(Qt.AlignCenter)
        self.crop_scroll = QScrollArea()
        self.crop_scroll.setWidget(self.crop_label)
        self.crop_scroll.setWidgetResizable(False)  # Не масштабировать!
        splitter.addWidget(self.crop_scroll)

        # Основной предпросмотр шаблона со вставкой в ScrollArea
        self.preview_label = QLabel("Появится шаблон для размещения вставки")
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_scroll = QScrollArea()
        self.preview_scroll.setWidget(self.preview_label)
        self.preview_scroll.setWidgetResizable(False)  # Не масштабировать!
        splitter.addWidget(self.preview_scroll)

        # Устанавливаем начальные пропорции (crop область побольше)
        splitter.setSizes([400, 300])
        layout.addWidget(splitter)

        # Slider для размера вставки (ширина, высота пропорциональна)
        slider_layout = QHBoxLayout()
        slider_label = QLabel("Размер вставки:")
        self.slider = QSlider(Qt.Horizontal)
        self.slider.setMinimum(50)
        self.slider.setMaximum(400)
        self.slider.setValue(150)
        self.slider.valueChanged.connect(self.change_insert_size)
        slider_layout.addWidget(slider_label)
        slider_layout.addWidget(self.slider)
        layout.addLayout(slider_layout)

        layout.addWidget(self.merge_btn)
        self.setLayout(layout)

    def choose_template(self):
        self.template_path, _ = QFileDialog.getOpenFileName(self, "Выбрать PDF шаблон")
        if not self.template_path:
            return
        self.status_label.setText("Шаблон выбран: "+self.template_path)
        doc = fitz.open(self.template_path)
        page = doc[0]
        pix = page.get_pixmap(dpi=150)
        self.preview_path = os.path.join(self.temp_dir, "template_preview.png")
        pix.save(self.preview_path)
        self.template_pixmap = QPixmap(self.preview_path)
        self.insert_pos = None
        self.show_template_preview()

    def choose_insert_pdf(self):
        self.insert_pdf_path, _ = QFileDialog.getOpenFileName(self, "Выбрать PDF вставку")
        if not self.insert_pdf_path:
            return
        self.status_label.setText("Вставка выбрана. Выдели область для вставки!")
        # Извлекаем первую страницу для crop
        insert_images = convert_from_path(self.insert_pdf_path, dpi=300, first_page=1, last_page=1)
        if insert_images:
            img = insert_images[0]
            self.sample_path = os.path.join(self.temp_dir, "sample_orig.png")
            img.save(self.sample_path)
            self.show_crop_preview()
        else:
            QMessageBox.warning(self, "Ошибка!", "Не удалось получить изображение из PDF!")
            return

    def show_crop_preview(self):
        self.crop_chosen = False
        pixmap = QPixmap(self.sample_path)
        self.crop_label.setPixmap(pixmap)
        self.crop_label.setScaledContents(False)  # Не масштабировать!
        self.crop_label.adjustSize()
        self.crop_label.mousePressEvent = self.start_crop_draw
        self.crop_label.mouseMoveEvent = self.crop_draw_rect
        self.crop_label.mouseReleaseEvent = self.finish_crop_draw

    def start_crop_draw(self, event):
        self.start_point = event.pos()
        self.end_point = event.pos()
        self.crop_chosen = False
        self.update_crop_preview()

    def crop_draw_rect(self, event):
        if self.start_point is None:
            return
        self.end_point = event.pos()
        self.update_crop_preview()

    def finish_crop_draw(self, event):
        self.end_point = event.pos()
        self.crop_chosen = True
        self.crop_rect = QRect(self.start_point, self.end_point)
        if self.crop_rect.width() > 10 and self.crop_rect.height() > 10:
            # Crop и показать дальше предпросмотр
            self.crop_png()
            self.status_label.setText("Область выбрана. Можно размещать вставку на шаблоне!")
            self.show_template_preview()
        else:
            self.status_label.setText("Слишком малая область для вставки!")
        self.update_crop_preview()

    def update_crop_preview(self):
        pixmap = QPixmap(self.sample_path)
        if self.start_point and self.end_point:
            painter = QPainter(pixmap)
            pen = QPen(Qt.green, 3, Qt.SolidLine)
            painter.setPen(pen)
            rect = QRect(self.start_point, self.end_point)
            painter.drawRect(rect)
            painter.end()
        self.crop_label.setPixmap(pixmap)
        self.crop_label.adjustSize()

    def crop_png(self):
        img = Image.open(self.sample_path)
        left = min(self.crop_rect.left(), self.crop_rect.right())
        upper = min(self.crop_rect.top(), self.crop_rect.bottom())
        right = max(self.crop_rect.left(), self.crop_rect.right())
        lower = max(self.crop_rect.top(), self.crop_rect.bottom())
        img_crop = img.crop((left, upper, right, lower))
        img_crop.save(self.crop_path)
        # Сохраняем пропорции crop
        crop_w = right - left
        crop_h = lower - upper
        self.insert_aspect_ratio = crop_w / crop_h if crop_h > 0 else 1.0
        # Инициализируем размеры вставки
        self.insert_width = 150
        self.insert_height = int(self.insert_width / self.insert_aspect_ratio)
        self.slider.setValue(self.insert_width)
        self.insert_images_ready = True

    def show_template_preview(self):
        if not os.path.exists(self.preview_path):
            return
        canvas_img = Image.open(self.preview_path).convert('RGBA')
        if os.path.exists(self.crop_path):
            insert_img = Image.open(self.crop_path).convert('RGBA')
            insert_img = insert_img.resize((self.insert_width, self.insert_height))
            # По центру или по последней позиции
            bg_w, bg_h = canvas_img.size
            if self.insert_pos is None:
                x = (bg_w - self.insert_width) // 2
                y = (bg_h - self.insert_height) // 2
                self.insert_pos = (x, y)
            x, y = self.insert_pos
            canvas_img.paste(insert_img, (x, y), insert_img)
        out_path = os.path.join(self.temp_dir, "preview_with_insert.png")
        canvas_img.save(out_path)
        pixmap = QPixmap(out_path)
        self.preview_label.setPixmap(pixmap)
        self.preview_label.setScaledContents(False)  # Не масштабировать!
        self.preview_label.adjustSize()
        # Mouse handlers для перетаскивания вставки
        self.preview_label.mousePressEvent = self.insert_press_event
        self.preview_label.mouseMoveEvent = self.insert_move_event
        self.preview_label.mouseReleaseEvent = self.insert_release_event

    def change_insert_size(self, value):
        self.insert_width = value
        self.insert_height = int(self.insert_width / self.insert_aspect_ratio)
        self.show_template_preview()

    def insert_press_event(self, event):
        if event.button() == Qt.LeftButton and self.insert_images_ready:
            self.dragging = True
            self.drag_offset = (event.pos().x() - self.insert_pos[0], event.pos().y() - self.insert_pos[1])

    def insert_move_event(self, event):
        if self.dragging and self.insert_images_ready:
            x = event.pos().x() - self.drag_offset[0]
            y = event.pos().y() - self.drag_offset[1]
            canvas_img = Image.open(self.preview_path).convert('RGBA')
            insert_img = Image.open(self.crop_path).convert('RGBA')
            insert_img = insert_img.resize((self.insert_width, self.insert_height))
            bg_w, bg_h = canvas_img.size
            # Границы
            x = max(0, min(bg_w - self.insert_width, x))
            y = max(0, min(bg_h - self.insert_height, y))
            self.insert_pos = (x, y)
            canvas_img.paste(insert_img, (x, y), insert_img)
            out_path = os.path.join(self.temp_dir, "preview_with_insert.png")
            canvas_img.save(out_path)
            pm = QPixmap(out_path)
            self.preview_label.setPixmap(pm)
            self.preview_label.adjustSize()

    def insert_release_event(self, event):
        self.dragging = False

    def safe_remove(self, path):
        for i in range(15):
            try:
                os.remove(path)
                return True
            except PermissionError:
                time.sleep(0.5)
        return False

    def start_merge(self):
        if not self.template_path or not self.insert_pdf_path or not self.crop_chosen or not self.insert_pos or not self.insert_images_ready:
            QMessageBox.warning(self, "Внимание", "Выбери шаблон, вставку, обрежь область и выбери позицию!")
            return
        self.status_label.setText("Извлекаю страницы и собираю PDF...")
        # 1. Извлечь все страницы из файла вставки
        insert_images = convert_from_path(self.insert_pdf_path, dpi=300)
        img_paths = []
        for i, img in enumerate(insert_images):
            fname = os.path.join(self.temp_dir, f"orig_{i:04d}.png")
            img.save(fname)
            img_paths.append(fname)
        # 2. Обрезать каждый файл аналогично выбранному crop
        cropped_paths = []
        left = min(self.crop_rect.left(), self.crop_rect.right())
        upper = min(self.crop_rect.top(), self.crop_rect.bottom())
        right = max(self.crop_rect.left(), self.crop_rect.right())
        lower = max(self.crop_rect.top(), self.crop_rect.bottom())
        for fname in img_paths:
            img = Image.open(fname)
            crop = img.crop((left, upper, right, lower))
            crop_path = os.path.join(self.temp_dir, f"crop_{os.path.basename(fname)}")
            crop.save(crop_path)
            cropped_paths.append(crop_path)
        # 3. Получить размеры шаблона в PDF-координатах
        doc_template = fitz.open(self.template_path)
        template_page = doc_template[0]
        pdf_width = template_page.rect.width
        pdf_height = template_page.rect.height
        pixmap_width = self.template_pixmap.width()
        pixmap_height = self.template_pixmap.height()
        # 4. Перевести позицию и размер вставки из предпросмотра в PDF-координаты
        x_scale = pdf_width / pixmap_width
        y_scale = pdf_height / pixmap_height
        insert_x = int(self.insert_pos[0] * x_scale)
        insert_y = int(self.insert_pos[1] * y_scale)
        insert_w = int(self.insert_width * x_scale)
        insert_h = int(self.insert_height * y_scale)
        if insert_w <= 0 or insert_h <= 0:
            QMessageBox.warning(self, "Ошибка!", "Размер вставки некорректен!")
            for fname in img_paths + cropped_paths:
                self.safe_remove(fname)
            return
        rect = fitz.Rect(insert_x, insert_y, insert_x + insert_w, insert_y + insert_h)
        output_pdf_name = "output.pdf"
        output_doc = fitz.open()
        for i in range(len(cropped_paths)):
            new_page = output_doc.new_page(width=pdf_width, height=pdf_height)
            new_page.show_pdf_page(template_page.rect, doc_template, 0)
            new_page.insert_image(rect, filename=cropped_paths[i])
        output_doc.save(output_pdf_name)
        # Очистить временные файлы
        for f in img_paths + cropped_paths:
            self.safe_remove(f)
        if os.path.exists(self.preview_path): self.safe_remove(self.preview_path)
        if os.path.exists(self.sample_path): self.safe_remove(self.sample_path)
        if os.path.exists(self.crop_path): self.safe_remove(self.crop_path)
        preview_with_insert = os.path.join(self.temp_dir, "preview_with_insert.png")
        if os.path.exists(preview_with_insert): self.safe_remove(preview_with_insert)
        self.status_label.setText(f"Готово! Новый файл: {output_pdf_name}")
        QMessageBox.information(self, "Готово!", f"Файл сформирован: {output_pdf_name}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFMergerApp()
    window.show()
    sys.exit(app.exec_())
