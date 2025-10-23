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
        self.setWindowTitle("PDF QR Merger")
        self.setMinimumWidth(900)
        self.setMinimumHeight(700)
        self.template_path = ""
        self.qr_pdf_path = ""
        self.preview_path = ""
        self.qr_rect = None
        self.qr_area_chosen = False

        self.start_point = None
        self.end_point = None
        self.template_pixmap = None

        self.qr_crop_path = "qr_crop_preview.png"
        self.qr_sample_path = ""
        self.qr_crop_rect = None
        self.qr_crop_chosen = False

        self.qr_pos = None   # x, y pos on preview
        self.qr_width = 150  # Default QR width in px (on preview)
        self.qr_height = 150 # Default QR height in px (on preview)
        self.qr_aspect_ratio = 1.0  # width/height
        self.dragging = False
        self.slider = None   # QSlider for QR size

        self.qr_images_ready = False
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        self.template_btn = QPushButton("Выбрать шаблон PDF")
        self.qr_btn = QPushButton("Выбрать PDF с QR-кодами")
        self.merge_btn = QPushButton("Слияние")
        self.status_label = QLabel("1. Шаблон PDF. 2. PDF с QR. 3. Crop QR на белом, 4. Слияние.")

        self.template_btn.clicked.connect(self.choose_template)
        self.qr_btn.clicked.connect(self.choose_qr_pdf)
        self.merge_btn.clicked.connect(self.start_merge)

        layout.addWidget(self.template_btn)
        layout.addWidget(self.qr_btn)
        layout.addWidget(self.status_label)

        # Создаём QSplitter вертикальный
        splitter = QSplitter(Qt.Vertical)

        # Crop preview для QR в ScrollArea
        self.qr_crop_label = QLabel("Появится QR для crop")
        self.qr_crop_label.setAlignment(Qt.AlignCenter)
        self.qr_crop_scroll = QScrollArea()
        self.qr_crop_scroll.setWidget(self.qr_crop_label)
        self.qr_crop_scroll.setWidgetResizable(False)  # Не масштабировать!
        splitter.addWidget(self.qr_crop_scroll)

        # Основной предпросмотр шаблона c QR в ScrollArea
        self.preview_label = QLabel("Появится шаблон для размещения QR")
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_scroll = QScrollArea()
        self.preview_scroll.setWidget(self.preview_label)
        self.preview_scroll.setWidgetResizable(False)  # Не масштабировать!
        splitter.addWidget(self.preview_scroll)

        # Устанавливаем начальные пропорции (crop QR побольше)
        splitter.setSizes([400, 300])
        layout.addWidget(splitter)

        # Slider для размера QR (ширина, высота пропорциональна)
        slider_layout = QHBoxLayout()
        slider_label = QLabel("Размер QR:")
        self.slider = QSlider(Qt.Horizontal)
        self.slider.setMinimum(50)
        self.slider.setMaximum(400)
        self.slider.setValue(150)
        self.slider.valueChanged.connect(self.change_qr_size)
        slider_layout.addWidget(slider_label)
        slider_layout.addWidget(self.slider)
        layout.addLayout(slider_layout)

        layout.addWidget(self.merge_btn)
        self.setLayout(layout)

    def choose_template(self):
        self.template_path, _ = QFileDialog.getOpenFileName(self, "Выбрать шаблон PDF")
        if not self.template_path:
            return
        self.status_label.setText("Шаблон выбран: "+self.template_path)
        doc = fitz.open(self.template_path)
        page = doc[0]
        pix = page.get_pixmap(dpi=150)
        self.preview_path = "template_preview.png"
        pix.save(self.preview_path)
        self.template_pixmap = QPixmap(self.preview_path)
        self.qr_pos = None
        self.show_template_preview()

    def choose_qr_pdf(self):
        self.qr_pdf_path, _ = QFileDialog.getOpenFileName(self, "Выбрать PDF с QR-кодами")
        if not self.qr_pdf_path:
            return
        self.status_label.setText("QR выбран. Выдели область QR!")
        # Извлекаем первый QR-код для crop
        qr_images = convert_from_path(self.qr_pdf_path, dpi=300, first_page=1, last_page=1)
        if qr_images:
            img = qr_images[0]
            self.qr_sample_path = "qr_sample_orig.png"
            img.save(self.qr_sample_path)
            self.show_qr_crop_preview()
        else:
            QMessageBox.warning(self, "Ошибка!", "Не удалось получить QR изображение!")
            return

    def show_qr_crop_preview(self):
        self.qr_crop_chosen = False
        pixmap = QPixmap(self.qr_sample_path)
        self.qr_crop_label.setPixmap(pixmap)
        self.qr_crop_label.setScaledContents(False)  # Не масштабировать!
        self.qr_crop_label.adjustSize()
        self.qr_crop_label.mousePressEvent = self.start_crop_draw
        self.qr_crop_label.mouseMoveEvent = self.crop_draw_rect
        self.qr_crop_label.mouseReleaseEvent = self.finish_crop_draw

    def start_crop_draw(self, event):
        self.start_point = event.pos()
        self.end_point = event.pos()
        self.qr_crop_chosen = False
        self.update_crop_preview()

    def crop_draw_rect(self, event):
        if self.start_point is None:
            return
        self.end_point = event.pos()
        self.update_crop_preview()

    def finish_crop_draw(self, event):
        self.end_point = event.pos()
        self.qr_crop_chosen = True
        self.qr_crop_rect = QRect(self.start_point, self.end_point)
        if self.qr_crop_rect.width() > 10 and self.qr_crop_rect.height() > 10:
            # Crop QR и показать дальше предпросмотр
            self.crop_qr_png()
            self.status_label.setText("QR crop выбран. Можно размещать QR на шаблоне!")
            self.show_template_preview()
        else:
            self.status_label.setText("Слишком малая область для QR!")
        self.update_crop_preview()

    def update_crop_preview(self):
        pixmap = QPixmap(self.qr_sample_path)
        if self.start_point and self.end_point:
            painter = QPainter(pixmap)
            pen = QPen(Qt.green, 3, Qt.SolidLine)
            painter.setPen(pen)
            rect = QRect(self.start_point, self.end_point)
            painter.drawRect(rect)
            painter.end()
        self.qr_crop_label.setPixmap(pixmap)
        self.qr_crop_label.adjustSize()

    def crop_qr_png(self):
        img = Image.open(self.qr_sample_path)
        left = min(self.qr_crop_rect.left(), self.qr_crop_rect.right())
        upper = min(self.qr_crop_rect.top(), self.qr_crop_rect.bottom())
        right = max(self.qr_crop_rect.left(), self.qr_crop_rect.right())
        lower = max(self.qr_crop_rect.top(), self.qr_crop_rect.bottom())
        img_crop = img.crop((left, upper, right, lower))
        img_crop.save(self.qr_crop_path)
        # Сохраняем пропорции crop
        crop_w = right - left
        crop_h = lower - upper
        self.qr_aspect_ratio = crop_w / crop_h if crop_h > 0 else 1.0
        # Инициализируем размеры QR
        self.qr_width = 150
        self.qr_height = int(self.qr_width / self.qr_aspect_ratio)
        self.slider.setValue(self.qr_width)
        self.qr_images_ready = True

    def show_template_preview(self):
        if not os.path.exists(self.preview_path):
            return
        canvas_img = Image.open(self.preview_path).convert('RGBA')
        if os.path.exists(self.qr_crop_path):
            qr_img = Image.open(self.qr_crop_path).convert('RGBA')
            qr_img = qr_img.resize((self.qr_width, self.qr_height))
            # По центру или по последней позиции
            bg_w, bg_h = canvas_img.size
            if self.qr_pos is None:
                x = (bg_w - self.qr_width) // 2
                y = (bg_h - self.qr_height) // 2
                self.qr_pos = (x, y)
            x, y = self.qr_pos
            canvas_img.paste(qr_img, (x, y), qr_img)
        out_path = "preview_with_realqr.png"
        canvas_img.save(out_path)
        pixmap = QPixmap(out_path)
        self.preview_label.setPixmap(pixmap)
        self.preview_label.setScaledContents(False)  # Не масштабировать!
        self.preview_label.adjustSize()
        # Mouse handlers для перетаскивания QR
        self.preview_label.mousePressEvent = self.qr_press_event
        self.preview_label.mouseMoveEvent = self.qr_move_event
        self.preview_label.mouseReleaseEvent = self.qr_release_event

    def change_qr_size(self, value):
        self.qr_width = value
        self.qr_height = int(self.qr_width / self.qr_aspect_ratio)
        self.show_template_preview()

    def qr_press_event(self, event):
        if event.button() == Qt.LeftButton and self.qr_images_ready:
            self.dragging = True
            self.drag_offset = (event.pos().x() - self.qr_pos[0], event.pos().y() - self.qr_pos[1])

    def qr_move_event(self, event):
        if self.dragging and self.qr_images_ready:
            x = event.pos().x() - self.drag_offset[0]
            y = event.pos().y() - self.drag_offset[1]
            canvas_img = Image.open(self.preview_path).convert('RGBA')
            qr_img = Image.open(self.qr_crop_path).convert('RGBA')
            qr_img = qr_img.resize((self.qr_width, self.qr_height))
            bg_w, bg_h = canvas_img.size
            # Границы
            x = max(0, min(bg_w - self.qr_width, x))
            y = max(0, min(bg_h - self.qr_height, y))
            self.qr_pos = (x, y)
            canvas_img.paste(qr_img, (x, y), qr_img)
            out_path = "preview_with_realqr.png"
            canvas_img.save(out_path)
            pm = QPixmap(out_path)
            self.preview_label.setPixmap(pm)
            self.preview_label.adjustSize()

    def qr_release_event(self, event):
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
        if not self.template_path or not self.qr_pdf_path or not self.qr_crop_chosen or not self.qr_pos or not self.qr_images_ready:
            QMessageBox.warning(self, "Внимание", "Выбери шаблон, QR, обрежь QR и выбери его позицию!")
            return
        self.status_label.setText("Извлекаю QR и собираю PDF...")
        # 1. Извлечь все QR из файла
        qr_images = convert_from_path(self.qr_pdf_path, dpi=300)
        img_paths = []
        for i, img in enumerate(qr_images):
            fname = f"qrorig_{i:04d}.png"
            img.save(fname)
            img_paths.append(fname)
        # 2. Обрезать каждый QR-файл аналогично выбранному crop
        qr_cropped_paths = []
        left = min(self.qr_crop_rect.left(), self.qr_crop_rect.right())
        upper = min(self.qr_crop_rect.top(), self.qr_crop_rect.bottom())
        right = max(self.qr_crop_rect.left(), self.qr_crop_rect.right())
        lower = max(self.qr_crop_rect.top(), self.qr_crop_rect.bottom())
        for fname in img_paths:
            img = Image.open(fname)
            crop = img.crop((left, upper, right, lower))
            crop_path = f"qr_{os.path.basename(fname)}"
            crop.save(crop_path)
            qr_cropped_paths.append(crop_path)
        # 3. Получить размеры шаблона в PDF-координатах
        doc_template = fitz.open(self.template_path)
        template_page = doc_template[0]
        pdf_width = template_page.rect.width
        pdf_height = template_page.rect.height
        pixmap_width = self.template_pixmap.width()
        pixmap_height = self.template_pixmap.height()
        # 4. Перевести позицию и размер QR из предпросмотра в PDF-координаты
        x_scale = pdf_width / pixmap_width
        y_scale = pdf_height / pixmap_height
        qr_x = int(self.qr_pos[0] * x_scale)
        qr_y = int(self.qr_pos[1] * y_scale)
        qr_w = int(self.qr_width * x_scale)
        qr_h = int(self.qr_height * y_scale)
        if qr_w <= 0 or qr_h <= 0:
            QMessageBox.warning(self, "Ошибка!", "Размер QR некорректен!")
            for fname in img_paths + qr_cropped_paths:
                self.safe_remove(fname)
            return
        rect = fitz.Rect(qr_x, qr_y, qr_x + qr_w, qr_y + qr_h)
        output_pdf_name = "output.pdf"
        output_doc = fitz.open()
        for i in range(len(qr_cropped_paths)):
            new_page = output_doc.new_page(width=pdf_width, height=pdf_height)
            new_page.show_pdf_page(template_page.rect, doc_template, 0)
            new_page.insert_image(rect, filename=qr_cropped_paths[i])
        output_doc.save(output_pdf_name)
        # Очистить временные файлы
        for f in img_paths + qr_cropped_paths:
            self.safe_remove(f)
        if os.path.exists(self.preview_path): self.safe_remove(self.preview_path)
        if os.path.exists(self.qr_sample_path): self.safe_remove(self.qr_sample_path)
        if os.path.exists(self.qr_crop_path): self.safe_remove(self.qr_crop_path)
        if os.path.exists("preview_with_realqr.png"): self.safe_remove("preview_with_realqr.png")
        self.status_label.setText(f"Готово! Новый файл: {output_pdf_name}")
        QMessageBox.information(self, "Готово!", f"Файл сформирован: {output_pdf_name}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFMergerApp()
    window.show()
    sys.exit(app.exec_())
