# ARHerb

[English](#english) · [Русский](#русский)

## English

ARHerb is a desktop application that automates the creation of labels for herbarium specimens. It turns an Excel spreadsheet into a ready-to-print Word document, optionally with QR codes linking to the collection site on Google Maps.

### Features

- **Simple GUI** built with PyQt6.
- **Print-ready Word documents** (.docx), two labels per row; a label is never split across pages.
- **QR codes** with a Google Maps link built from the latitude/longitude columns.
- **Specimen types**: fungi, bryophytes and vascular plants (vascular plants can also show Russian family and species names).
- **Custom column mapping** if your spreadsheet uses different column names. Missing columns are reported before the document is generated.

### Download

A ready-made Windows build (`ARHerb.exe`) is available on the [Releases](https://github.com/sshskrbna/ARHerb/releases) page. The app tells you when a newer release is out.

### Run from source

Requires Python 3.10+.

```bash
git clone https://github.com/sshskrbna/ARHerb.git
cd ARHerb
pip install -r requirements.txt
python main.py
```

### Excel format

The first sheet is read; the first row holds column names. Default names:

| Column | Content |
|---|---|
| `family` | Family (Latin) |
| `familyrus` | Family (Russian), vascular plants only |
| `species` | Species (Latin) |
| `speciesrus` | Species (Russian), vascular plants only |
| `region` | Region |
| `point` | Locality |
| `habitats` | Habitat |
| `date` | Collection date |
| `leg.` | Collector |
| `det.` | Identified by |
| `num` | Specimen number |
| `N` | Latitude, decimal degrees (`55.75` or `55,75`) |
| `E` | Longitude, decimal degrees |

Empty cells are allowed. The output is saved next to the source file as `<name>_output.docx`.

### Tech stack

Python, PyQt6, pandas + openpyxl, python-docx, qrcode + Pillow.

### Contributing

Found a bug or have an idea? Please open an [Issue](https://github.com/sshskrbna/ARHerb/issues).

## Русский

ARHerb — десктопное приложение для автоматического создания этикеток гербарных образцов. Оно превращает Excel-таблицу в готовый к печати документ Word, при желании с QR-кодами, ведущими на место сбора в Google Maps.

### Возможности

- **Простой интерфейс** на PyQt6.
- **Готовые к печати документы Word** (.docx), по две этикетки в ряд; этикетка не разрывается между страницами.
- **QR-коды** со ссылкой на Google Maps по колонкам широты и долготы.
- **Типы образцов**: грибы, мохообразные и сосудистые растения (для сосудистых можно выводить русские названия семейства и вида).
- **Свои названия колонок**, если в таблице они отличаются от стандартных. Отсутствующие колонки показываются до генерации документа.

### Скачать

Готовая сборка для Windows (`ARHerb.exe`) лежит на странице [Releases](https://github.com/sshskrbna/ARHerb/releases). Приложение само сообщает о выходе новой версии.

### Запуск из исходников

Нужен Python 3.10+.

```bash
git clone https://github.com/sshskrbna/ARHerb.git
cd ARHerb
pip install -r requirements.txt
python main.py
```

### Формат Excel

Читается первый лист, в первой строке — названия колонок. Названия по умолчанию:

| Колонка | Содержимое |
|---|---|
| `family` | Семейство (латынь) |
| `familyrus` | Семейство (рус.), только для сосудистых |
| `species` | Вид (латынь) |
| `speciesrus` | Вид (рус.), только для сосудистых |
| `region` | Регион |
| `point` | Пункт сбора |
| `habitats` | Местообитание |
| `date` | Дата сбора |
| `leg.` | Собрал |
| `det.` | Определил |
| `num` | Номер образца |
| `N` | Широта в десятичных градусах (`55.75` или `55,75`) |
| `E` | Долгота в десятичных градусах |

Пустые ячейки допускаются. Результат сохраняется рядом с исходным файлом как `<имя>_output.docx`.

### Технологии

Python, PyQt6, pandas + openpyxl, python-docx, qrcode + Pillow.

### Участие в разработке

Нашли ошибку или есть идея? Создайте [Issue](https://github.com/sshskrbna/ARHerb/issues).
