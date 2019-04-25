Attribute VB_Name = "modGlobalDefinition"
Option Explicit
            'Объявить определяемый пользователем тип данных - строка "Системной таблицы"
Type SystemInfo
            'Имя (Пароля, Контроллера, Ридера, ... и т.д.)
    strObject As String * 16
            'Константа или Адрес и Терминал
    strConsAddrTerm As String * 16
            'Тип (Константа, Ридер, Райтер, Препроцессор)
    strType As String * 2
            'Индекс - ссылка на номер строки "Системной таблицы"
    strIndex As String * 6
            'Резерв
    strAppendix As String * 8
End Type
            'Объявить определяемый пользователем тип данных - строка "Таблицы персон"
Type PersonInfo
            'Фамилия (имя)
    strName As String * 16
            'Персональный код
    strPersonCode As String * 16
            'Статус
    strStatus As String * 2
            'Время
    strTime As String * 8
            'Календарь
    strCalendar As String * 2
            'Резерв
    strReserve As String * 8
End Type


            'Объявить переменную-объект для его связывания
            '   c ActiveX.EXE
Global objTablePerson As Sel_2Server.XTablePerson
            'Объявить переменную-объект для его связывания только
            '   c объяленными интерфейсами "FlexGrid"
            '   ("Таблицы персон")
Global gTablePerson As Sel_2Server.ITablePerson

'ОТЛАДКА
'Global objTablePerson As XTablePerson
'Global gTablePerson As ITablePerson
            
            'Объявить переменную-объект MSMQQueueInfo -
            '   необходим для создания и управления
            '   очередью ПРИНИМАЕМЫХ СООБЩЕНИЙ
Global qInfoInput As MSMQQueueInfo
            'Переменная-объект ОЧЕРЕДЬ ПРИНИМАЕМЫХ СООБЩЕНИЙ
Global qQueueInput As MSMQQueue
            'Переменная-объект ОЧЕРЕДЬ-СОБЫТИЕ
            ' ПРИНИМАЕМЫХ СООБЩЕНИЙ
Global evQueue As MSMQQueue
            'Переменная-объект ПРИНИМАЕМОЕ СООБЩЕНИE
Global qMsgInput As MSMQMessage
            'Строка ПРИНИМАЕМЫХ метки и текста СООБЩЕНИЯ
Global strMsgInput As String
            'Объявить переменную-объект MSMQQueueInfo -
            '   необходим для управления
            '   очередью ПЕРЕДАВАЕМЫХ СООБЩЕНИЙ
Global qInfoOutput As MSMQQueueInfo
            'Переменная-объект ОЧЕРЕДЬ ПЕРЕДАВАЕМЫХ СООБЩЕНИЙ
Global qQueueOutput As MSMQQueue
            'Переменная-объект ПЕРЕДАВАЕМОЕ СООБЩЕНИE
Global qMsgOutput As MSMQMessage
            
            'Объявить определяемый пользователем тип данных - строка "Таблицы календаря"
Type CalendarInfo
            'Номер недели
    strWeekNum As String * 20
            'Понедельник - дата
    strMonday As String * 4
            'Вторник - дата
    strTuesday As String * 4
            'Среда - дата
    strWednesday As String * 4
            'Четверг - дата
    strThursday As String * 4
            'Пятница - дата
    strFriday As String * 4
            'Суббота - дата
    strSaturday As String * 4
            'Воскресенье
    strSunday As String * 4
End Type

            'Объявить определяемый пользователем тип данных - строка "Таблицы времени"
Type TimeInfo
            'Номер интервала времени
    strIntervalNum As String * 12
            'Время
    strTime As String * 8
            'Расширение
    strExpander As String * 8
End Type
            
            'Объявить определяемый пользователем тип данных - строка "Таблицы терминалов"
Type TerminalInfo
            'Терминал (Имя из ячейки "Object" "Системной таблицы")
    strTerminal As String * 16
            'Адрес и Порт
    strAddrPort As String * 4
            'Описание терминала
    strDescription As String * 16
            'Расширение
    strExpander As String * 8
End Type

            'Объявить определяемый пользователем тип данных - строка "Таблицы протокола"
Type ProtocolInfo
            'Фамилия (имя) или терминал
    strProtocName As String * 16
            'Персональный код или пароль
    strProtocPersonCode As String * 16
            'Статус
    strProtocStatus As String * 22
            'Время
    strProtocTime As String * 10
            'Дата
    strProtocDate As String * 10
            'Резерв
    strProtocReserve As String * 22
End Type
            
            'Строка "Таблицы протокола"
Global gProtocol As ProtocolInfo
            'Флаг запрета опроса терминалов
Global gTermContr As Integer
            'Флаг Шины Управления
Global gBus As Integer
            'Массив адресов "Controller'ов" для Ручного Управления
            '  электронной "Кнопкой"
Global gAddrManual(3) As String * 2
            'Опция "Шлюз" для терминалов порта
Global gSluice(3) As Integer
            'Начальный символ в Имени персоны - Признак Гостя (т.е. не Служащего)
Global gVisitor As String
            'Флаг запрета автоматической перезаписи "Таблицы протокола" в Базу
Global gMSBase As Integer
            'Типы устройств для "Печати Документа" (1 - Простой Чековый принтер,
            '  2 - Принтер ШтрихКода, 4 - Кассовый Чековый принтер;
            '  Возможны комбинации: 0 - нет устройств, 1+2, 1+4, 2+4, 1+2+4)
Global gDocument As Integer
            'Печать - (Название Компании)
Global gPrintSIAName As String * 16
            'Длина (Размер) въездного/входного талона
            '  до линии отреза/отрыва
Global gTalonLength As Integer
            'Опция автоматического формирования полей "PersonCode" и "Info"
            '  0 - Ручное формирование; 1 - Автоматическое для Предприятия;
            '  2 - Автоматическое для Автостоянки. Возможна комбинация: 1+2
Global gCreatePersonCode As Integer
            'Дата и время формирования последнего "Z_Отчета"
Global gZ_Report As String
            'Индекс въездного (входного) терминала
Global gTermInp As Integer
            'Индекс выездного (выходного) терминала
Global gTermOut As Integer
            
            'Умалчиваемый временной интервал при АвтоРегистрации работников
Global gDefaultTime As String
            'Умалчиваемое значение ячейки "Status"
            '  в "Таблице персон" при АвтоРегистрации работников
Global gDefaultStatus As String
            'Умалчиваемое значение ячейки "Calendar"
            '  в "Таблице персон" при АвтоРегистрации работников
Global gDefaultCalendar As String
            'Признак необходимости автоматического формирования
            '  опций выходных дней в "Таблице календаря" для
            '  Нового Года
Global gHolidays As Integer
            'Массив адресов "Controller'ов" для Подтверждения
            ' электронной "Кнопкой" входа (выхода) Служащих
Global gEmplAddrTerm(3) As String * 2
            
            
            'Тариф пользования Автостоянкой (Предприятием) для Специальных
            '  Клиентов (время пользования не учитывается - константная оплата)
Global gTariffConst As Integer
            
            'Умалчиваемое значение ячейки "Calendar"
            '  в "Таблице персон" при АвтоРегистрации на Автостоянке
Global gDefaultParkCale As String
            'Массив адресов "Controller'ов" Автостоянки для Подтверждения
            ' электронной "Кнопкой" въезда (выезда) Клиентов
Global gParkAddrTerm(3) As String * 2
            'Дневной тариф Автостоянки (для постоянных Клиентов)
Global gParkingD As Integer
            'Ночной тариф Автостоянки (для постоянных Клиентов)
Global gParkingN As Integer
            'Суточный тариф Автостоянки (для постоянных Клиентов)
Global gParkingDN As Integer
            'Часовой Дневной тариф Автостоянки (для временных Клиентов)
Global gParkingHourD As Integer
            'Часовой Ночной тариф Автостоянки (для временных Клиентов)
Global gParkingHourN As Integer
            'Дневное время допуска к Автостоянке (для постояннных Клиентов)
Global gParkingTimeD As String
            'Умалчиваемое время допуска к Автостоянке (для Всех Клиентов)
Global gDefaultParkTime As String
            'Опция дискретности (точности) учета времени парковки
            '  при Регистрации/Исключении Временного Клиента Автостоянки
Global gParkingTimeCell As Integer
            'Количество ячеек времени, в течение которого разрешается
            '  АМ Постоянного Клиента непрерывно находиться на Автостоянке
            '  (Переменная вычисляется каждый раз при старте Приложения через
            '   переменные gParkTimeLimit и gParkingTimeCell)
Global gParkingCellLimit As Integer
            'Количество оплачиваемых предварительно ячеек времени
            '  при Регистрации Временных Клиентов
Global gParkInpCellNumb As Integer
            'Въездной тариф Автостоянки (для временных Клиентов)
Global gParkingMoneyCell As Integer
            'Опция "Физическое/Логическое удаление" при Исключении
            '  Клиента Автостоянки из "Таблицы персон"
Global gParkingDeletion As Integer
            'Опция копирования "PersonCode"в поле "Info"
            '  при Регистрации Временного Клиента Автостоянки
Global gParkingCodeInfo As Integer
            'Опция ручного подтверждения открытия шлагбаума
            '  при Регистрации/Исключении Временного Клиента Автостоянки
Global gParkingPresButton As Integer
            'Лимит (в мин.) времени непрерывного пребывания
            '  Постоянного Клиента на Автостоянке
Global gParkTimeLimit As Integer
            'Количество мест на Автостоянке
Global gParkingPlaceNum As Integer
            'Количество свободных мест на Автостоянке
Global gParkFreePlaces As Integer

            'Умалчиваемое значение ячейки "Calendar"
            '  в "Таблице персон" при АвтоРегистрации Посетителей
Global gDefaultAcceCale As String
            'Массив адресов "Controller'ов" Предприятия для Подтверждения
            ' электронной "Кнопкой" входа (выхода) Посетителей
Global gAcceAddrTerm(3) As String * 2
            'Дневной тариф Предприятия (для постоянных Клиентов)
Global gAccessD As Integer
            'Ночной тариф Предприятия (для постоянных Клиентов)
Global gAccessN As Integer
            'Суточный тариф Предприятия (для постоянных Клиентов)
Global gAccessDN As Integer
            'Часовой Дневной тариф Предприятия (для временных Клиентов)
Global gAccessHourD As Integer
            'Часовой Ночной тариф Предприятия (для временных Клиентов)
Global gAccessHourN As Integer
            'Дневное время допуска к Предприятия (для постояннных Клиентов)
Global gAccessTimeD As String
            'Умалчиваемое время допуска на Предприятие (для Всех Клиентов)
Global gDefaultAcceTime As String
            'Опция дискретности (точности) учета времени посещения
            '  при Регистрации/Исключении Временного Клиента
Global gAccessTimeCell As Integer
            '  Количество ячеек времени, в течение которого разрешается
            '   Постоянному Посетителю непрерывно находиться на Предприятии
            '  (Переменная вычисляется каждый раз при старте Приложения через
            '   переменные gAcceTimeLimit и gAccessTimeCell)
Global gAccessCellLimit As Integer
            'Количество оплачиваемых предварительно ячеек времени
            '  при Регистрации Временных Клиентов
Global gAcceInpCellNumb As Integer
            'Входной тариф Предприятия (для временных Клиентов)
Global gAccessMoneyCell As Integer
            'Входной тариф Предприятия для Взрослых (для временных Клиентов)
Global gAccessMoneyCellHuman As Integer
            'Входной тариф Предприятия для Детей (для временных Клиентов)
Global gAccessMoneyCellBaby As Integer
            'Входной тариф Предприятия для Конвоя (для временных Клиентов)
Global gAccessMoneyCellConvoy As Integer
            'Входной тариф Предприятия для Семьи (для временных Клиентов)
Global gAccessMoneyCellFamily As Integer
            'Опция "Физическое/Логическое удаление" при Исключении
            '  Клиента Предприятия из "Таблицы персон"
Global gAccessDeletion As Integer
            'Опция копирования "PersonCode"в поле "Info"
            '  при Регистрации Временного Клиента Предприятия
Global gAccessCodeInfo As Integer
            'Опция ручного подтверждения открытия турникета
            '  при Регистрации/Исключении Временного Клиента Предприятия
Global gAccessPresButton As Integer
            'Стоимость ПРОКАТА ИНВЕНТАРЯ # 1, 2, 3 и 4
Global gLease1 As Integer
Global gLease2 As Integer
Global gLease3 As Integer
Global gLease4 As Integer
            'Лимит (в мин.) времени непрерывного пребывания
            '  Постоянного Клиента на Предприятии
Global gAcceTimeLimit As Integer
            'Количество мест на Предприятии
Global gAccessPlaceNum As Integer
            'Количество свободных мест на Предприятии
Global gAcceFreePlaces As Integer
            
            'Полное имя файла "Архив протокола" (с указанием "пути" к нему)
Global gPathFileName As String
            'Ячейка из "Таблицы календаря" - (Текущий день)
Global gToday() As String * 4
            'Номер строки "Таблицы календаря", где расположена ячейка Текущего дня
Global gRowNum As Integer
            'Номер столбца "Таблицы календаря", где расположена ячейка Текущего дня
Global gColNum As Integer
            'Массив интервалов доступа для всех вариантов "Таблицы времени"
Global gInterval() As String * 11
            'Массив дополнительных терминалов и календарей для
            '  всех вариантов "Таблицы времени"
Global gTerCal() As String * 12
            'Массив терминалов доступа для всех вариантов "Таблицы терминалов"
Global gAddrPort() As String * 4
            'Текущий номер  свободной строки "Таблицы протокола"
Global gProtocRowNum As Integer
            'Текущий номер файла "Таблицы протокола"
Global gProtocFileNum As Integer
            'Текущий номер  свободной строки DUMMY файла
Global gDummyRowNum As Long
            'Текущий номер DUMMY файла
Global gFileDummy As Integer
            'Имя Загрузочного модуля (Для процедуры StartUp)
Global gModuleStartUp As String
            'Текущий номер копируемого в Препроцессор файла
Global gPreprocFileNum As Integer
            'Имя Препроцессора в локальной сети
Global gPreprocName As String
            'Имя "Host'a" в локальной сети
Global gHost As String
            'Количество Процессоров в локальной сети (исключая
            '  собственный Препроцессор)
Global gNetPreprocNum As Integer
            'Массив Имен Процессоров локальной сети (исключая
            '  собственный Препроцессор)
Global gSocketNet() As String
            'Индекс Препроцессора (номер строки "Системной таблицы"
            '   Препроцессора с информацией о нем)
Global gPreprocIndex As Integer
            'Длина персонального кода для входного терминала J-го порта
Global gPersonCode(3) As Integer
            'Количество удалений/добавлений строк в "Системной таблице"
Global gAddDelRowTableSystem
            'Признак внесения изменений в "Системную таблицу"
Global gChangesTableSystem As Boolean
            'Признак внесения изменений в "Таблицу персон"
Global gChangesTablePerson As Boolean
            'Флаг реального (физического) удаления строк из "Таблицы персон"
Global gRealDelPerson As Boolean
            'Признак внесения изменений в "Таблицу календарь"
Global gChangesTableCalendar As Boolean
            'Количество вариантов "Таблицы календаря"
Global gVarNumCalendar As Integer
            'Количество удалений/добавлений строк в "Таблице времени"
Global gAddDelRowTableTime
            'Признак внесения изменений в "Таблицу времени"
Global gChangesTableTime As Boolean
            'Количество вариантов "Таблицы времени"
Global gVarNumTime As Integer
            'Количество удалений/добавлений строк в "Таблице терминалов"
Global gAddDelRowTableTerminal
            'Признак внесения изменений в "Таблицу терминалов"
Global gChangesTableTerminal As Boolean
            'Количество вариантов "Таблицы терминалов"
Global gVarNumTerminal As Integer
            'Признак  внесения изменений в "Таблицу температур"
Global gChangesTableTemperature As Boolean
            'Признак  внесения изменений в "Таблицу мощностей"
Global gChangesTablePower As Boolean
             'Количество строк на одной странице формы "frmPrintPreview"
Global gRowPrintQuan As Integer
            'Количество дней (обратный отсчет, начиная с текущего дня),
            '  которые просматриваются системой при копировании
            '  Архивов Препроцессоа в "Host Computer" и при
            '  формировании из Архивов Баз Данных
Global gDayNum As Integer
            'Признак необходимости сжатия "Таблицы персон":
            '   устанавливается всегда в "Host Computer'e" и в тех случаях
            '   в "Препроцессоре", когда последний использует свою
            '   собственную "Таблицу персон" - "ЗЕРКАЛЬНАЯ Таблицa персон"
Global gCompresTablPers As Integer
            'Текущий год
Global gYear As Integer
            'Опция разделения времени (параллельное выполнение процедур)
Global gTimeShare As Integer
            'Длительность звукового сигнала
Global gBeepSound As Integer
            'Минимальный возможный номер при автоматическом
            '  формировании поля "PersonCode"
Global gMinCount As Integer
            'Максимальный возможный номер при автоматическом
            '  формировании поля "PersonCode"
Global gMaxCount As Integer
            'Уменьшение показаний количества свободных мест
            '  на дисплее по сравнению со счетчиком свободных мест
            '  (для устранения конфликтов при нескольких
            '  входах/въездах)
Global gDisplayDiscount As Integer

            'Декларация API функций, необходимых для работы
            '  с объектом "Mutex"
Declare Function WaitForSingleObject Lib "kernel32" ( _
   ByVal hHandle As Long, _
   ByVal dwMilliseconds As Long) As Long

Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" ( _
   ByVal lpMutexAttributes As Long, _
   ByVal bInitialOwner As Long, _
   ByVal lpName As String) As Long

Declare Function ReleaseMutex Lib "kernel32" ( _
   ByVal hMutex As Long) As Long
   
            'Дескриптор объекта "Mutex", необходимого для
            '  синхронизации потоков, совместно использующих
            '  общие ресурсы - атрибут доступности "Таблицы Персон"
            '  и саму "Таблицу Персон" (при вычеркивании ее строк)
Global gMutex As Long


