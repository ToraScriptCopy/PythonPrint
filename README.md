# PythonPrint

## Основной функционал

* **Direct Print (GDI)**: рендеринг и отправка данных напрямую в контекст устройства (Device Context). Исключает вызов сторонних ассоциативных приложений.
* **Native PDF Rendering**: интеграция библиотеки MuPDF для растеризации PDF-документов перед отправкой в очередь печати.
* **Win32 Spooler Integration**: управление параметрами DEVMODE (копии, выбор лотка, формат носителя) через системные вызовы Windows.
* **DND Interface**: поддержка протокола Drag-and-Drop для формирования очереди задач.
* **Live Preview**: динамическая генерация эскизов на основе первой страницы документа для верификации перед печатью.
* **Checklist Management**: выборочная активация/деактивация объектов в пуле задач.

## Стек технологий

* **Frontend**: CustomTkinter (UI Engine), tkinterdnd2 (DND Extension).
* **Backend**: pywin32 (Win32Print, Win32UI, Win32Con).
* **Processing**: PyMuPDF (PDF Engine), Pillow (Image Processing).

## Использование

При запуске инициализируется модуль `bootstrap`, который автоматически проверяет и устанавливает отсутствующие зависимости.

```bash
python main.py
