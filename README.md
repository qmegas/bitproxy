# BitTorrent Proxy

Written using Visual Basic 5

Project freezed since 2011

# Release history (russian only)

BitTorrent Proxy v1.46 (27.06.2011)
-----------------------------------
- Мелкие косметические доработки
- Исправлены множественные ошибки


BitTorrent Proxy v1.45 (10.06.2011)
-----------------------------------
- Изменять имя задачи эмуляции и скачки можно прям в списке не
  переходя в окно редактирования
- Новая опция в настроках эмуляции: Разрешать добавлять торренты с 
  одинаковым хешем"
- Исправлен критический баг
- Исправлены орфографические ошибки в языковых модулях


BitTorrent Proxy v1.44 (30.04.2011)
-----------------------------------
- В списках скачки и эмуляции теперь сохраняются размеры и порядок 
  колонок
- В эмуляции задачи с установленым автостопом выделяются синим цветом
- Ошибка при добавлении в список эмуляции торрента который уже там 
  имеется отображается теперь как Tray Ballon, а не как сообщение.
- При двойном щелчке на нижнюю панель в главном окне, сообщение в 
  этой панеле удаляется
- Исправлена ошибка при включенной опции "Не отправлять данные о 
  количестве скаченого"
- Исправлена ошибка при которой трекеру передавались не верные данные 
  когда задача в эмуляции останавливалась, а потом запускалась заново
- Исправлены найденные орфографические ошибки в Русском и Украинском 
  языках
- Новая колонка "Добавлено" в списке эмуляций
- Настройки автоостановки сбрасываются после срабатывания
- Новый режим автоостановки после определенного времени
- Файл клиента BitComet 1.22 был заменен на версию 1.27
- Файл клиента uTorrent 2.0.3 был заменен на версию 2.2


BitTorrent Proxy v1.43 (11.02.2011)
-----------------------------------
- Исправлена ошибка при востановлении иконки в трее
- Исправлена ошибка при которой не сохранялись параметры постепенного
  увеличения скорости в эмуляции
- Из настроек убрана опция "Определять добавленные торренты, как уже 
  100% скаченные", теперь этот параметр сохраняется как и все 
  остальные параметры при добавлении торрента в эмуляцию


BitTorrent Proxy v1.42 (22.01.2011)
-----------------------------------
- В случае падения эксплорера, программа востанавливает значек в 
  системном трее
- В эмуляции теперь можно редактировать название раздачи
- При добавлении в эмуляцию нескольких торрентов одновременно, есть 
  опция установить одни и те же настройки для всех торрентов
- Удалена тестовая опция "Сохранять настройки для каждого трекера 
  отдельно"


BitTorrent Proxy v1.41 RC2 (05.11.2010)
-----------------------------------
- Исправлена ошибка при которой запрос к трекеру создавался с ошибкой
- Исправлена ошибка при которой программа зависала если обрабатывала 
  не правильно созданный торрент файл
- Добавлена проверка уже существующих торрент файлов в списке эмуляции
- Создана отдельная страница для загрузки новых файлов торрент клиентов
- Добавлена возможность скачать файл клиента uTorrent 1.6.1


BitTorrent Proxy v1.41 RC1 (20.08.2010)
-----------------------------------
- Начало работы над изменением файлов клиентов.


BitTorrent Proxy v1.40 (21.05.2010)
-----------------------------------
- Улучшена система защиты от одного пира
- Серийный ключ больше не нужен. Полная версия стала доступна всем.


BitTorrent Proxy v1.36 (20.03.2010)
-----------------------------------
- Файл клиента uTorrent обновился до версии 2.0
- Мелкие баг-фиксы


BitTorrent Proxy v1.35 (14.12.2009)
-----------------------------------
- Исправление критической ошибки в режиме скачки


BitTorrent Proxy v1.34 (11.12.2009)
-----------------------------------
- Небольшая перестановка в окне настроек
- Новая опция: не удалять ретракер из списка анонса
- Исправлена ошибка при которой программа не могла работать с торрент
  файлами в которых небыло значения announce
- Исправлена ошибка при которой нельзя было выключить компьютер
  если программа запущена
- Файл клиента BitComet обновлен до версии 1.16
- Удаленное управление отключено по умолчанию


BitTorrent Proxy v1.33 (24.10.2009)
-----------------------------------
- Добавлена поддержка защиты трекера, когда трекер возвращает HTTP код 
  302 - редайрект
- Добавлен немецкий перевод программы
- Мелкие визуальные исправления


BitTorrent Proxy v1.32 (19.09.2009)
-----------------------------------
- Исправлена ошибка в режиме скачки которая иногда появлялась при 
  включении опции смены битторрент клиента
- В инсталяционный пакет включены исправленные файлы клиентов uTorrent.
  Также файл клиента uTorrent обновлен до версии 1.84
- Мелкие дополнительные оптимизации в полной версии программы


BitTorrent Proxy v1.31 (17.08.2009)
-----------------------------------
- Функция сброса скорости до нуля в случае если трекер возвращает 
  только одного пира
- Исправлен баг при котором в режиме скачки не сохранялась опция
  "Уменьшать загрузку в..."
- Исправления мелких багов
- Файл клиента BitComet обновлен до версии 1.14


BitTorrent Proxy v1.30 (31.07.2009)
-----------------------------------
- Полная передлка диалога режима скачки
- Убран режим HTTP Proxy
- Добавлен график отображающий изменение в скорости при раздаче, то как
  это видит трекер
- Добавлена опция "не отсылать количество скачиваемого" в режиме скачки
- Добавлена система плавного изменения скорости при скачивании
- Добавлена возможность изменять скорость сразу на нескольких задач
  эмуляции
- Начата реализация удаленного управления программой
- Языковые модули теперь представляют из себя обычные текстовые файлы
- Исправлена ошибка в эмуляции из-за которой иногда программе не 
  удавалось правильно определить время обновления передаваемое трекером
- Исправлена ошибка при обработки торрентов со списком анонса из 
  нескольких трекеров
- Сохранение списков эмуляции и скачки теперь происходит при любом
  изменении этих списков, а не при выходе из программы
- В полной версии передана система активации программы
- Обновлены файлы клиентов. Добавлены клиенты: BitComet 1.13 и 
  uTorrent 1.83


BitTorrent Proxy v1.22 (04.04.2009)
-----------------------------------
* Версия только для обладателей полной версией программы
- Оптимизация работы с битторрент клиентом uTorrent
- В окне скачки, при режиме динамического изменения, в логах пишется 
  расчетная скорость отдачи
- Оптимизация кода связанная с модулями перевода. Теперь программа 
  запускается, даже если в папке нет ниодного языкового модуля.
- Опция "Игнорировать ошибки соединения" в эмуляции
- Опция остоновки эмуляции в случае если программа не может подключится 
  к серверу длитильное время
- Опция грубой остановки эмуляции, без обновления данных на трекере
- Мелкие доработки внешнего вида программы
- Файл клиента uTorrent обновлен до версии 1.8.2
- Файл клиента BitComet обновлен до версии 1.10


BitTorrent Proxy v1.21 (27.12.2008)
-----------------------------------
- Исправлена ошибка с неправильным подсчетом байт в автоостановке
- Исправлена ошибка с невозможностью изменить значение в поле "порт"
- Исправлена ошибка не позволяющая открывать файлы содержащие символы 
  юникода в названии файла или в его пути
- Добавлено всплывающее меню при щелчке правой кнопкой на иконке в 
  системном трее
- Файл клиента BitComet обновлен до версии 1.07
- Добавлен немецкий язык в инсталяционный пакет


BitTorrent Proxy v1.20 (30.10.2008)
-----------------------------------
- Появилась возможность сортировать список задачи эмуляции
- Небольшие косметические доработки
- Главное окно и окно эмуляции теперь запоминают свой размер
- Новая опция в эмуляции "Постепенное наращивание скорости"
- При ошибках в эмуляции теперь возле иконки трея выскакивает сообщение
- Новая опция в настройках "Игнорировать ошибки сервера при эмуляции"
- Исправлена ошибка, когда изза таймаута останавливалась задача в 
  эмуляции


BitTorrent Proxy v1.19 (13.08.2008)
-----------------------------------
- Появилась возможность перетаскивать торрент файлы в окно программы из 
  проводника Windows
- Также можно перетаскивать файлы в окно эмуляции, но тогда файлы 
  попадут именно в эмуляцию
- Добавлены две горячие клавиши для эмуляции: Ctrl-A - выделяет все 
  элементы списка, Delete - удаляет элементы списка
- Новая опция в настройках: "Определять по умолчанию, добавленные 
  торренты, как уже 100% скаченные"
- Файл клиента uTorrent обновлен до версии 1.8


BitTorrent Proxy v1.18 (20.07.2008)
-----------------------------------
- Теперь при включенной опции сохранения списка заданий, также 
  автоматически сохраняется список заданий сгрузки
- Новая опция в настройках: Сворачиваться в трей при закрытии окна
- Новая опция в настройках: Автоматически загружать окна закачек при 
  запуске программы
- Новая опция в настройках: Запускаться при запуске Windows
- Мелкие исправления в модуле англиского перевода


BitTorrent Proxy v1.17 (17.05.2008)
-----------------------------------
- Поддержка русских имен при обработке торрента
- Теперь порт в эмуляции имеет всегда одно и тоже значение
- В эмуляции возможность запускать, удалять и останавливать сразу 
  несколько задач


BitTorrent Proxy v1.16 (24.03.2008)
-----------------------------------
- При удалении задания из эмуляции, появляется подтверждающее диалоговое 
  окно
- Добавлена опция "Добавить к даунлоаду" в эмуляции
- Новая кнопка "Помощь" в главном окне
- В скачке, режим одиночного и нескольких подключений переименованы в 
  режимы динамического и постоянного изменения


BitTorrent Proxy v1.15 (26.10.2007)
-----------------------------------
- Можно сохранять и востанавливать окона загрузок
- Опция автоостановки в эмуляции. Пока два режима
* Останавливать эмуляцию в зависимости от количества загруженного
* Останавливать эмуляцию в зависимости от рейтинга
- В эмуляции при ручном добавлении к аплоаду можно выбирать единицы: 
  Kb, Mb, Gb, Tb
- Возможность получать количество пиров и сидов при эмуляции
- Мелкие дополнительные исправления


BitTorrent Proxy v1.1 (06.10.2007)
-----------------------------------
- При запуске скачки через системное меню в окне скачки отображается 
  название раздачи
- Оптимизирована работа с торрент файлами при эмуляции
- При возвращении трекером ошибки при эмуляции, отображается причина 
  ошибки
- Исправлена ошибка возникающая при не возможности подключится к трекеру 
  во время эмуляции
- Некоторые внутренние оптимизации


BitTorrent Proxy v1.0 (12.09.2007)
-----------------------------------
- Использование одного окна для всех копий программы
- Максимальное ограниче коэффицента при скачки: 9999999
- Возможность уменьшать количество сгрузки
- Переделка системы поддержки разных языков
- Переделка системы эмулируемых битторрент клиентов
* Новый режим эмуляции
- опция игнора трекерного времени обновления
- Опция запуска задачи эмуляции с определённого значения
- Опция ручного добавления аплоад трафика к эмуляции
- сохранение списка эмуляций
- Перенос настроек в отдельное окно
- Опция выбора действия по умолчанию
- Опция "порт по умолчанию"
- Опция отключения сохранения списка эмулируемых задач
- Кнопка "Сгрузка файлов клиентов"


BitTorrent Proxy v0.89 (18.07.2007)
-----------------------------------
* Версия только для обладателей полной версией программы
- Слежение за статусом раздачи при использовании режима одиночного 
  подключения


BitTorrent Proxy v0.88 (19.06.2007)
-----------------------------------
- починака бага в режиме "HTTP прокси", когда трекер использовал для 
  подключения какой-то не стандартный порт
- исправления бага с announce-list
- максимальное увеличение в полной версии в 999 раз


BitTorrent Proxy v0.87 (20.05.2007)
-----------------------------------
- множественные багфиксы


BitTorrent Proxy v0.86 (09.05.2007)
-----------------------------------
* Версия для тестеров
- возможность использовать программу как обычный HTTP прокси
- поддержка внешнего прокси


BitTorrent Proxy v0.85 (07.05.2007)
-----------------------------------
- поддержка нескольких языков
- добавлен Англиский язык
- инсталяция для программы
- множественные багфиксы


BitTorrent Proxy v0.84 (25.04.2007)
-----------------------------------
- возможность отключать втоматическую проверку обновлений
- оптимизация работы с uTorrent
- возможность работы с трекерами использующие нестандартный порт 
  подключения


BitTorrent Proxy v0.83 (21.04.2007)
-----------------------------------
- улучшена система передачи данных закодированных с помощью gzip
- починка многих багов


BitTorrent Proxy v0.82 (09.04.2007)
-----------------------------------
- Первая публичная версия
- Автовыбор свободного порта
- Сохранение последних 10 трекеров
- Оптимизация работы с uTorrent


BitTorrent Proxy v0.8 (03.04.2007)
----------------------------------
- Пункт в системном меню
- Автозагрузка всех данных торрента
- Проверка новых версий
- Ручная опция проверки новых версий


BitTorrent Proxy v0.71 (01.03.2007)
-----------------------------------
- Возможность менять трекер не перезапуская программу
- Некоторые исправления в коде