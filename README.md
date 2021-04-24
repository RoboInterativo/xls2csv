# xls2csv
```
├── config.yaml
├── in
│   └── ODMC.xlsx
├── in_arc
├── LICENSE
├── out
│   └── ODMC.csv
├── README.md
├── requirements.txt
└── xls2csv.py
```
Для использования активируйте виртуальную среду.

```
source venv/bin/activate
```

При запуске конвертирует все файлы из папки  in

Кладет csv в out

Исходный файл переносит в архив in_arc

для использования ./xls2csv.py

csv разделитель запятая, кодировка UTF-8

Возможные проблемы:

Если не работает из предложенного окружения, пересоздайте его 

```
rm -rf venv
mkdir venv
virtualenv venv
```

Активируйте окружение и установите зависимости запустив

``` 
source venv/bin/activate
./install.sh
```


TODO config.yaml для изменения конфигурации
