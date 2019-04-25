Attribute VB_Name = "HostApp"
Option Explicit
            'Опция "Логического Выключения" терминалов - при неисправностях и пр.
Dim intTerminalLogOFF As Integer
            'Признак "Белый ключ" (ПЕРСОНАЛЬНЫЙ КОД хранится в
            '   локальной памяти 'Controllera"
Dim intWhite As Integer
            'Опция "Автостоянка" для терминалов порта
Dim intParking(3) As Integer
            'Опция "Посетитель" для терминалов порта
Dim intAccess(3) As Integer
            'Опция "Служащий" для терминалов порта
Dim intEmploye(3) As Integer
            'Индекс порта на форме Главного окна Приложения
Dim intIndex As Integer
            'Адрес "Controller'a" и "Port'a" для АвтоРегистрации/Удаления
Dim intAutoRegDel(3) As Integer
            'Опция  "Пароль операторов"
Dim strParole As String
            'Опция "Печать Документа" при Регистрации/Удаления Клиента
Dim intDocument As Integer
            'Время - Часы
Dim intHour As Integer
Dim strHour As String
            'Время - Минуты
Dim intMinute As Integer
Dim strMinute As String
            'Номер текущей строки в "Таблице персон"
Dim intRowNum As Integer
            'Индекс элемента в массиве элементов управления форм
Dim intControlIndex As Integer
            'Буфер приема данных от "Controller'ов" терминалов
Dim vntBufferInput(3) As Variant
            'Строка отсылаемого сообщения
Dim strMessage As String

            ' Главная процедура - обработка событий ввода/вывода для терминалов
Public Sub Main()
            'Номер текущей строки в "Системной таблице"
Dim intRowNumSys As Integer
            'Номер позиции признака "/" в анализируемом поле
Dim intPosNum As Integer
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Рабочий счетчик
Dim intCount As Integer
            'Рабочeе полe
Dim intWork As Integer

            ' Загрузить не показывая форму "frmTableSystem"
    Load frmTableSystem
            ' Загрузить не показывая форму "frmDemo"
    Load frmDemo
            'Сделать невидимым меню "Parking", "Access" и "Employe" формы "frmDemo"
    frmDemo.mnuParking.Visible = False
    frmDemo.mnuAccess.Visible = False
    frmDemo.mnuEmploye.Visible = False
            'Сделать недоступным меню "Parking", "Access" и "Employe" формы "frmDemo"
    frmDemo.mnuParking.Enabled = False
    frmDemo.mnuAccess.Enabled = False
    frmDemo.mnuEmploye.Enabled = False
    
            'Цикл по всем нефиксированным строкам "Системной таблицы"
    For intRowNumSys = 1 To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            'Текущая строка  "Системной таблицы"
        frmTableSystem.grdTableSystem.Row = intRowNumSys
            'Фиксированный столбец "Системной таблицы" (Объект)
        frmTableSystem.grdTableSystem.Col = 0
            'Инициализация констант Системы
        If Trim(frmTableSystem.grdTableSystem.Text) = "prtPortC(0)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить номер порта
            frmDemo.prtPortC(0).CommPort = Left(frmTableSystem.grdTableSystem.Text, 1)
            'Установить параметры порта
            intPosNum = InStr(3, frmTableSystem.grdTableSystem, "/")
            frmDemo.prtPortC(0).Settings = Mid(frmTableSystem.grdTableSystem.Text, 3, intPosNum - 3)
            'Установить свойство "InputMode" порта
            frmDemo.prtPortC(0).InputMode = Mid(frmTableSystem.grdTableSystem.Text, _
            intPosNum + 1, 1)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "prtPortC(1)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить номер порта
            frmDemo.prtPortC(1).CommPort = Left(frmTableSystem.grdTableSystem.Text, 1)
            'Установить параметры порта
            intPosNum = InStr(3, frmTableSystem.grdTableSystem, "/")
            frmDemo.prtPortC(1).Settings = Mid(frmTableSystem.grdTableSystem.Text, 3, intPosNum - 3)
            'Установить свойство "InputMode" порта
            frmDemo.prtPortC(1).InputMode = Mid(frmTableSystem.grdTableSystem.Text, _
            intPosNum + 1, 1)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "prtPortC(2)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить номер порта
            frmDemo.prtPortC(2).CommPort = Left(frmTableSystem.grdTableSystem.Text, 1)
            'Установить параметры порта
            intPosNum = InStr(3, frmTableSystem.grdTableSystem, "/")
            frmDemo.prtPortC(2).Settings = Mid(frmTableSystem.grdTableSystem.Text, 3, intPosNum - 3)
            'Установить свойство "InputMode" порта
            frmDemo.prtPortC(2).InputMode = Mid(frmTableSystem.grdTableSystem.Text, _
            intPosNum + 1, 1)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "prtPortC(3)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить номер порта
            frmDemo.prtPortC(3).CommPort = Left(frmTableSystem.grdTableSystem.Text, 1)
            'Установить параметры порта
            intPosNum = InStr(3, frmTableSystem.grdTableSystem, "/")
            frmDemo.prtPortC(3).Settings = Mid(frmTableSystem.grdTableSystem.Text, 3, intPosNum - 3)
            'Установить свойство "InputMode" порта
            frmDemo.prtPortC(3).InputMode = Mid(frmTableSystem.grdTableSystem.Text, _
            intPosNum + 1, 1)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ModuleStartUp" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Имя "StartUp" модуля системы
            gModuleStartUp = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "txtPassword" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Пароль системы
            frmDemo.txtPassword.Tag = Trim(frmTableSystem.grdTableSystem.Text)
            'Текущий столбец "Системной таблицы"=4(Резерв)
            frmTableSystem.grdTableSystem.Col = 4
            'Установить Псевдоним Препроцессора в локальной сети или
            '   "пусто" для "Host Computer'a"
            If Trim(frmTableSystem.grdTableSystem.Text) <> "" Then
                gPreprocName = Trim(frmTableSystem.grdTableSystem.Text)
            Else
                gPreprocName = ""
            'Резервировать нулевой элемент массива Имен Процессоров
            '  локальной сети
                ReDim gSocketNet(0) As String
                gSocketNet(0) = ""
            End If
            'Это не "Host Computer"
            If gPreprocName <> "" Then
            'Текущий столбец "Системной таблицы"=3(Индекс)
                frmTableSystem.grdTableSystem.Col = 3
            'Установить индекс Препроцессора (номер строки "Системной таблицы"
            '   Препроцессора с информацией о нем)
                gPreprocIndex = Trim(frmTableSystem.grdTableSystem.Text)
            'Резервировать нулевой элемент массива Имен Процессоров
            '  локальной сети
                ReDim gSocketNet(0) As String
                gSocketNet(0) = ""
            End If
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Host" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Имя "Host Computer'a" в локальной сети
            gHost = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "txtParole" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Пароль операторов
            strParole = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Visitor" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Признак Гостя
            gVisitor = Left(Trim(frmTableSystem.grdTableSystem.Text), 1)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Z_Report" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Дата и время формирования последнего "Z_Отчета"
            gZ_Report = Trim(frmTableSystem.grdTableSystem.Text)
        
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "TariffConst" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Тариф пользования Автостоянкой (Предприятием) для Специальных
            '  Клиентов (время пользования не учитывается - константная оплата)
            gTariffConst = Trim(frmTableSystem.grdTableSystem.Text)
            If gTariffConst > 32000 Then gTariffConst = 32000
        
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingPresButt" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Опцию ручного подтверждения открытия шлагбаума
            '   при Регистрации/Исключении Временных Клиентов Автостоянки
            gParkingPresButton = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingDeletion" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Опцию "Физическое/Логическое удаление"
            '   при Исключении Клиента Автостоянки из "Таблицы персон"
            gParkingDeletion = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingCode_Info" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Опцию копирования "PersonCode"в поле "Info"
            '   при Регистрации Временного Клиента Автостоянки
            gParkingCodeInfo = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingDN" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Суточный тариф автостоянки (для Постоянных Клиентов)
            gParkingDN = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingD" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Дневной тариф автостоянки (для Постоянных Клиентов)
            gParkingD = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingN" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Ночной тариф автостоянки (для Постоянных Клиентов)
            gParkingN = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingHourD" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Часовой Дневной тариф автостоянки (для Временных Клиентов)
            gParkingHourD = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingHourN" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Часовой Ночной тариф автостоянки (для Временных Клиентов)
            gParkingHourN = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingMoneyCell" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Въездной тариф автостоянки (для Временных Клиентов)
            gParkingMoneyCell = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingTimeCell" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить дискретность (точность) учета времени парковки
            '  при Регистрации/Исключении Временного Клиента Автостоянки
            gParkingTimeCell = Trim(frmTableSystem.grdTableSystem.Text)
            If gParkingTimeCell = 0 Then gParkingTimeCell = 15
            If gParkingTimeCell > 1440 Then gParkingTimeCell = 1440
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingTimeLimit" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить лимит (в сек.) времени непрерывного пребывания
            '  Постоянного Клиента на Автостоянке
            gParkTimeLimit = Trim(frmTableSystem.grdTableSystem.Text)
            If gParkTimeLimit > 1440 Then gParkTimeLimit = 1440
             'Количество ячеек времени, в течение которого разрешается
            '  АМ Постоянного Клиента непрерывно находиться на Автостоянке
            gParkingCellLimit = Int(gParkTimeLimit / gParkingTimeCell)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkInpCellNumb" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Количество оплачиваемых предварительно ячеек времени
            '  при Регистрации Временных Клиентов
            gParkInpCellNumb = Trim(frmTableSystem.grdTableSystem.Text)
            If gParkInpCellNumb * gParkingTimeCell > 1440 Then _
            gParkInpCellNumb = 1440 / gParkingTimeCell
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingPlaceNum" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Количество мест на Автостоянке
            gParkingPlaceNum = Trim(frmTableSystem.grdTableSystem.Text)
            If gParkingPlaceNum > 999 Then gParkingPlaceNum = 999
            'Количество свободных мест на Автостоянке
            gParkFreePlaces = gParkingPlaceNum
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessPresButt" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Опцию ручного подтверждения открытия турникета
            '   при Регистрации/Исключении Временных Посетителей Предприятия
            gAccessPresButton = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessDeletion" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Опцию "Физическое/Логическое удаление"
            '   при Исключении Посетителя Предприятия из "Таблицы персон"
            gAccessDeletion = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessCode_Info" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Опцию копирования "PersonCode"в поле "Info"
            '   при Регистрации Временного Посетителя Предприятия
            gAccessCodeInfo = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessDN" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Суточный тариф Предприятия (для Постоянных Посетителей)
            gAccessDN = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessD" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Дневной тариф Предприятия (для Постоянных Посетителей)
            gAccessD = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessN" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Ночной тариф Предприятия (для Постоянных Посетителей)
            gAccessN = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessHourD" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Часовой Дневной тариф Предприятия (для Временных Посетителей)
            gAccessHourD = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessHourN" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Часовой Ночной тариф Предприятия (для Временных Посетителей)
            gAccessHourN = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessMoneyCell" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Входной тариф для Взрослых (для Временных Посетителей)
            intPosNum = InStr(2, frmTableSystem.grdTableSystem, "/")
            gAccessMoneyCellHuman = _
            Left(frmTableSystem.grdTableSystem.Text, intPosNum - 1)
            gAccessMoneyCell = gAccessMoneyCellHuman
            'Установить Входной тариф для Детей (для Временных Посетителей)
            gAccessMoneyCellBaby = _
            Mid(frmTableSystem.grdTableSystem.Text, intPosNum + 1, _
            InStr(intPosNum + 1, frmTableSystem.grdTableSystem, "/") - intPosNum - 1)
            'Установить Входной тариф для Конвоя (для Временных Посетителей)
            intPosNum = InStr(intPosNum + 1, frmTableSystem.grdTableSystem, "/")
            gAccessMoneyCellConvoy = _
            Mid(frmTableSystem.grdTableSystem.Text, intPosNum + 1, _
            InStr(intPosNum + 1, frmTableSystem.grdTableSystem, "/") - intPosNum - 1)
            'Установить Входной тариф для Семьи (для Временных Посетителей)
            intPosNum = InStr(intPosNum + 1, frmTableSystem.grdTableSystem, "/")
            gAccessMoneyCellFamily = _
            Mid(frmTableSystem.grdTableSystem.Text, intPosNum + 1)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessTimeCell" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить дискретность (точность) учета времени посещения
            '  при Регистрации/Исключении Временного Посетителя Предприятия
            gAccessTimeCell = Trim(frmTableSystem.grdTableSystem.Text)
            If gAccessTimeCell = 0 Then gAccessTimeCell = 15
            If gAccessTimeCell > 1440 Then gAccessTimeCell = 1440
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessTimeLimit" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить лимит (в мин.) времени непрерывного пребывания
            '  Постоянного Клиента на Предприятии
            gAcceTimeLimit = Trim(frmTableSystem.grdTableSystem.Text)
            If gAcceTimeLimit > 1440 Then gAcceTimeLimit = 1440
             'Количество ячеек времени, в течение которого разрешается
            '  Постоянному Клиенту непрерывно находиться на Предприятии
            '  и количество предоплаченных ячеек времени для Временного Клиента
            gAccessCellLimit = Int(gAcceTimeLimit / gAccessTimeCell)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AcceInpCellNumb" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Количество оплачиваемых предварительно ячеек времени
            '  при Регистрации Временных Клиентов
            gAcceInpCellNumb = Trim(frmTableSystem.grdTableSystem.Text)
            If gAcceInpCellNumb * gAccessTimeCell > 1440 Then _
            gAcceInpCellNumb = 1440 / gAccessTimeCell
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessPlaceNum" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Количество мест на Предприятии
            gAccessPlaceNum = Trim(frmTableSystem.grdTableSystem.Text)
            If gAccessPlaceNum > 999 Then gAccessPlaceNum = 999
            'Количество свободных мест на Предприятии
            gAcceFreePlaces = gAccessPlaceNum
            
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Document" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Опция "Печать Документа" при Регистрации/Удаления Клиента
            gDocument = Trim(frmTableSystem.grdTableSystem.Text)
            If gDocument <> 0 Then intDocument = 1
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "prtDocument" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить номер порта для простого чекового принтера
            frmDemo.prtPortDocument.CommPort = Left(frmTableSystem.grdTableSystem.Text, 1)
            'Установить параметры порта для простого чекового принтера
            intPosNum = InStr(3, frmTableSystem.grdTableSystem, "/")
            frmDemo.prtPortDocument.Settings = Mid(frmTableSystem.grdTableSystem.Text, 3, intPosNum - 3)
            'Установить свойство "InputMode" порта для простого чекового принтера
            frmDemo.prtPortDocument.InputMode = Mid(frmTableSystem.grdTableSystem.Text, _
            intPosNum + 1, 1)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "prtBarCode" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить номер порта для принтера штрих-кода
            frmDemo.prtPortBarCode.CommPort = Left(frmTableSystem.grdTableSystem.Text, 1)
            'Установить параметры порта для принтера штрих-кода
            intPosNum = InStr(3, frmTableSystem.grdTableSystem, "/")
            frmDemo.prtPortBarCode.Settings = Mid(frmTableSystem.grdTableSystem.Text, 3, intPosNum - 3)
            'Установить свойство "InputMode" порта для принтера штрих-кода
            frmDemo.prtPortBarCode.InputMode = Mid(frmTableSystem.grdTableSystem.Text, _
            intPosNum + 1, 1)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "prtDisplay" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить номер порта для дисплея количества свободных мест
            frmDemo.prtPortDisplay.CommPort = Left(frmTableSystem.grdTableSystem.Text, 1)
            'Установить параметры порта для дисплея количества свободных мест
            intPosNum = InStr(3, frmTableSystem.grdTableSystem, "/")
            frmDemo.prtPortDisplay.Settings = Mid(frmTableSystem.grdTableSystem.Text, 3, intPosNum - 3)
            'Установить свойство "InputMode" порта для дисплея количества
            '  свободных мест
            frmDemo.prtPortDisplay.InputMode = Mid(frmTableSystem.grdTableSystem.Text, _
            intPosNum + 1, 1)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Lease" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить стоимость ПРОКАТА ИНВЕНТАРЯ # 1
            intPosNum = InStr(2, frmTableSystem.grdTableSystem, "/")
            gLease1 = Left(frmTableSystem.grdTableSystem.Text, intPosNum - 1)
            'Установить стоимость ПРОКАТА ИНВЕНТАРЯ # 2
            gLease2 = Mid(frmTableSystem.grdTableSystem.Text, intPosNum + 1, _
            InStr(intPosNum + 1, frmTableSystem.grdTableSystem, "/") - intPosNum - 1)
            'Установить стоимость ПРОКАТА ИНВЕНТАРЯ # 3
            intPosNum = InStr(intPosNum + 1, frmTableSystem.grdTableSystem, "/")
            gLease3 = Mid(frmTableSystem.grdTableSystem.Text, intPosNum + 1, _
            InStr(intPosNum + 1, frmTableSystem.grdTableSystem, "/") - intPosNum - 1)
            'Установить стоимость ПРОКАТА ИНВЕНТАРЯ # 4
            intPosNum = InStr(intPosNum + 1, frmTableSystem.grdTableSystem, "/")
            gLease4 = Mid(frmTableSystem.grdTableSystem.Text, intPosNum + 1)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "tmrTimeOut(0)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить время "TimeOut' для "Controller'ов" N_0
            frmDemo.tmrTimeOut(0).Interval = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "tmrTimeOut(1)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить время "TimeOut' для "Controller'ов" N_1
            frmDemo.tmrTimeOut(1).Interval = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "tmrTimeOut(2)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить время "TimeOut' для "Controller'ов" N_2
            frmDemo.tmrTimeOut(2).Interval = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "tmrTimeOut(3)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить время "TimeOut' для "Controller'ов" N_3
            frmDemo.tmrTimeOut(3).Interval = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "tmrButton(0)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить время "TimeOut' для "Кнопки подтверждения" N_0
            frmDemo.tmrButton(0).Interval = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "tmrButton(1)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить время "TimeOut' для "Кнопки подтверждения" N_1
            frmDemo.tmrButton(1).Interval = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "tmrButton(2)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить время "TimeOut' для "Кнопки подтверждения" N_2
            frmDemo.tmrButton(2).Interval = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "tmrButton(3)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить время "TimeOut' для "Кнопки подтверждения" N_3
            frmDemo.tmrButton(3).Interval = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "tmrTermContr" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить интервал опроса "Controller'ов" терминалов
            frmDemo.tmrTermContr.Interval = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "tmrPasswTimeOut" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить время "TimeOut' для ввода пароля
            frmDemo.tmrPasswTimeOut.Interval = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "chkTerm(0)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить значение флага обслуживания терминалов N_0
            frmDemo.chkTerm(0).Value = frmTableSystem.grdTableSystem.Text
            'Признак "неустановленного" индекса въездного (входного) терминала
            gTermInp = -1
            'Признак "неустановленного" индекса выездного (выходного) терминала
            gTermOut = -1
            'Если терминалы N_0 обслуживаются системой
            If frmDemo.chkTerm(0).Value = 1 Then
            'Текущий столбец "Системной таблицы"=4(Резерв)
                frmTableSystem.grdTableSystem.Col = 4
            'Если это въездной (входной) терминал системы
                If Trim(frmTableSystem.grdTableSystem.Text) = "Inp" Then
            'Установить индекс въездного (входного) терминала
                    gTermInp = 0
            'Если это выездной (выходной) терминал системы
                ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Out" Then
            'Установить индекс выездного (выходного) терминала
                    gTermOut = 0
            'Если это въездной/выездной (входной/выходной) терминал системы
                ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "InpOut" Then
            'Установить индекс въездного (входного) терминала
                    gTermInp = 0
            'Установить индекс выездного (выходного) терминала
                    gTermOut = 0
                End If
            End If
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "chkTerm(1)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить значение флага обслуживания терминалов N_1
            frmDemo.chkTerm(1).Value = frmTableSystem.grdTableSystem.Text
            'Если терминалы N_1 обслуживаются системой
            If frmDemo.chkTerm(1).Value = 1 Then
            'Текущий столбец "Системной таблицы"=4(Резерв)
                frmTableSystem.grdTableSystem.Col = 4
            'Если это въездной (входной) терминал системы
                If Trim(frmTableSystem.grdTableSystem.Text) = "Inp" Then
            'Установить "неустановленный" индекс въездного (входного) терминала
                    If gTermInp = -1 Then gTermInp = 1
            'Если это выездной (выходной) терминал системы
                ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Out" Then
            'Установить "неустановленный" индекс выездного (выходного) терминала
                    If gTermOut = -1 Then gTermOut = 1
            'Если это въездной/выездной (входной/выходной) терминал системы
                ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "InpOut" Then
            'Установить "неустановленный" индекс въездного (входного) терминала
                    If gTermInp = -1 Then gTermInp = 1
            'Установить "неустановленный" индекс выездного (выходного) терминала
                    If gTermOut = -1 Then gTermOut = 1
                End If
            End If
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "chkTerm(2)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить значение флага обслуживания терминалов N_2
            frmDemo.chkTerm(2).Value = frmTableSystem.grdTableSystem.Text
            'Если терминалы N_2 обслуживаются системой
            If frmDemo.chkTerm(2).Value = 1 Then
            'Текущий столбец "Системной таблицы"=4(Резерв)
                frmTableSystem.grdTableSystem.Col = 4
            'Если это въездной (входной) терминал системы
                If Trim(frmTableSystem.grdTableSystem.Text) = "Inp" Then
            'Установить "неустановленный" индекс въездного (входного) терминала
                    If gTermInp = -1 Then gTermInp = 2
            'Если это выездной (выходной) терминал системы
                ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Out" Then
            'Установить "неустановленный" индекс выездного (выходного) терминала
                    If gTermOut = -1 Then gTermOut = 2
            'Если это въездной/выездной (входной/выходной) терминал системы
                ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "InpOut" Then
            'Установить "неустановленный" индекс въездного (входного) терминала
                    If gTermInp = -1 Then gTermInp = 2
            'Установить "неустановленный" индекс выездного (выходного) терминала
                    If gTermOut = -1 Then gTermOut = 2
                End If
            End If
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "chkTerm(3)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить значение флага обслуживания терминалов N_3
            frmDemo.chkTerm(3).Value = frmTableSystem.grdTableSystem.Text
            'Если терминалы N_3 обслуживаются системой
            If frmDemo.chkTerm(3).Value = 1 Then
            'Текущий столбец "Системной таблицы"=4(Резерв)
                frmTableSystem.grdTableSystem.Col = 4
            'Если это въездной (входной) терминал системы
                If Trim(frmTableSystem.grdTableSystem.Text) = "Inp" Then
            'Установить "неустановленный" индекс въездного (входного) терминала
                    If gTermInp = -1 Then gTermInp = 3
            'Если это выездной (выходной) терминал системы
                ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Out" Then
            'Установить "неустановленный" индекс выездного (выходного) терминала
                    If gTermOut = -1 Then gTermOut = 3
            'Если это въездной/выездной (входной/выходной) терминал системы
                ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "InpOut" Then
            'Установить "неустановленный" индекс въездного (входного) терминала
                    If gTermInp = -1 Then gTermInp = 3
            'Установить "неустановленный" индекс выездного (выходного) терминала
                    If gTermOut = -1 Then gTermOut = 3
                End If
            End If
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "chkPhoto(0)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить значение флага обслуживания фото для терминалов N_0
            frmDemo.chkPhoto(0).Value = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "chkPhoto(1)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить значение флага обслуживания фото для терминалов N_1
            frmDemo.chkPhoto(1).Value = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "chkPhoto(2)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить значение флага обслуживания фото для терминалов N_2
            frmDemo.chkPhoto(2).Value = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "chkPhoto(3)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить значение флага обслуживания фото для терминалов N_3
            frmDemo.chkPhoto(3).Value = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "optEnglish" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить значение опции для Английского языка
            frmDemo.optEnglish.Value = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "optLatvian" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить значение опции для Латышского языка
            frmDemo.optLatvian.Value = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "optRussian" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить значение опции для Русского языка
            frmDemo.optRussian.Value = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "optAutomatic" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить значение опции для Автоматического управления терминалами
            frmDemo.optAutomatic.Value = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "optManual" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить значение опции для Ручного управления терминалами
            frmDemo.optManual.Value = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gPersonCode(0)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить длину персонального кода для "Controller'ов" N_0
            gPersonCode(0) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gPersonCode(1)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить длину персонального кода для "Controller'ов" N_1
            gPersonCode(1) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gPersonCode(2)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить длину персонального кода для "Controller'ов" N_2
            gPersonCode(2) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gPersonCode(3)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить длину персонального кода для "Controller'ов" N_3
            gPersonCode(3) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gRowPrintQuan" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить количество строк на одной странице формы "frmPrintPreview"
            gRowPrintQuan = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gDayNum" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить количество дней (обратный отсчет, начиная с текущего дня),
            '  которые просматриваются системой при копировании
            '  Архивов Препроцессоа в "Host Computer" и при
            '  формировании из Архивов Баз Данных
            gDayNum = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gCompresTablPers" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить признак необходимости сжатия "Таблицы персон"
            gCompresTablPers = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gYear" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить текущий год
            gYear = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gVarNumTime" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить количество вариантов "Таблицы времени"
            gVarNumTime = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gVarNumCalendar" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить количество вариантов "Таблицы календаря"
            gVarNumCalendar = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gVarNumTerminal" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить количество вариантов "Таблицы терминалов"
            gVarNumTerminal = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AddrManual(0)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить адрес "Controller'a" для Ручного Управления
            '  электронной "Кнопкой" "N_0"
            gAddrManual(0) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AddrManual(1)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить адрес "Controller'a" для Ручного Управления
            '  электронной "Кнопкой" "N_1"
            gAddrManual(1) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AddrManual(2)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить адрес "Controller'a" для Ручного Управления
            '  электронной "Кнопкой" "N_2"
            gAddrManual(2) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AddrManual(3)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить адрес "Controller'a" для Ручного Управления
            '  электронной "Кнопкой" "N_3"
            gAddrManual(3) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AutoRegDel(0)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить aдрес "Controller'a" и "Port'a" для
            '  АвтоРегистрации/Удаления
            intAutoRegDel(0) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AutoRegDel(1)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить aдрес "Controller'a" и "Port'a" для
            '  АвтоРегистрации/Удаления
            intAutoRegDel(1) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AutoRegDel(2)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить aдрес "Controller'a" и "Port'a" для
            '  АвтоРегистрации/Удаления
            intAutoRegDel(2) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AutoRegDel(3)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить aдрес "Controller'a" и "Port'a" для
            '  АвтоРегистрации/Удаления
            intAutoRegDel(3) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Sluice(0)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "Шлюз" для терминалов "N_0" -
            '  полный цикл
            gSluice(0) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Sluice(1)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "Шлюз" для терминалов "N_1" -
            '  полный цикл
            gSluice(1) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Sluice(2)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "Шлюз" для терминалов "N_2" -
            '  полный цикл
            gSluice(2) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Sluice(3)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "Шлюз" для терминалов "N_3" -
            '  полный цикл
            gSluice(3) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "TerminalLogOFF" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "Логическое Выключение" терминалов -
            '  при неисправностях и пр.
            intTerminalLogOFF = frmTableSystem.grdTableSystem.Text
            
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Parking(0)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "Автостоянка" для терминалов "N_0"
            intParking(0) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Parking(1)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "Автостоянка" для терминалов "N_1"
            intParking(1) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Parking(2)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "Автостоянка" для терминалов "N_2"
            intParking(2) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Parking(3)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "Автостоянка" для терминалов "N_3"
            intParking(3) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkAddrTerm(0)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "АдресТерминалаАвтостоянки" для терминалов "N_0"
            gParkAddrTerm(0) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkAddrTerm(1)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "АдресТерминалаАвтостоянки" для терминалов "N_1"
            gParkAddrTerm(1) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkAddrTerm(2)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "АдресТерминалаАвтостоянки" для терминалов "N_2"
            gParkAddrTerm(2) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkAddrTerm(3)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "АдресТерминалаАвтостоянки" для терминалов "N_3"
            gParkAddrTerm(3) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingTimeD" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Дневной временной интервал автостоянки
            '   при АвтоРегистрации (для Постоянных Клиентов)
            gParkingTimeD = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "DefaultParkTime" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Умалчиваемый временной интервал доступа к
            '   Автостоянки при АвтоРегистрации Клиентов
            gDefaultParkTime = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "DefaultParkCale" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить умалчиваемое значение ячейки "Calendar"
            '  в "Таблице персон" при АвтоРегистрации Клиентов на Автостоянке
            gDefaultParkCale = Trim(frmTableSystem.grdTableSystem.Text)
            
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Access(0)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "Посетитель" для терминалов "N_0"
            intAccess(0) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Access(1)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "Посетитель" для терминалов "N_1"
            intAccess(1) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Access(2)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "Посетитель" для терминалов "N_2"
            intAccess(2) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Access(3)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "Посетитель" для терминалов "N_3"
            intAccess(3) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AcceAddrTerm(0)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "АдресТерминалаПосетителя" для терминалов "N_0"
            gAcceAddrTerm(0) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AcceAddrTerm(1)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "АдресТерминалаПосетителя" для терминалов "N_1"
            gAcceAddrTerm(1) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AcceAddrTerm(2)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "АдресТерминалаПосетителя" для терминалов "N_2"
            gAcceAddrTerm(2) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AcceAddrTerm(3)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "АдресТерминалаПосетителя" для терминалов "N_3"
            gAcceAddrTerm(3) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessTimeD" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Дневной временной интервал посещения
            '   при АвтоРегистрации (для Постоянных Клиентов)
            gAccessTimeD = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "DefaultAcceTime" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Умалчиваемый временной интервал доступа
            '   при АвтоРегистрации Посетителей
            gDefaultAcceTime = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "DefaultAcceCale" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить умалчиваемое значение ячейки "Calendar"
            '  в "Таблице персон" при АвтоРегистрации Посетителей
            gDefaultAcceCale = Trim(frmTableSystem.grdTableSystem.Text)
            
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Employe(0)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "Служащий" для терминалов "N_0"
            intEmploye(0) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Employe(1)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "Служащий" для терминалов "N_1"
            intEmploye(1) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Employe(2)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "Служащий" для терминалов "N_2"
            intEmploye(2) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Employe(3)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "Служащий" для терминалов "N_3"
            intEmploye(3) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "EmplAddrTerm(0)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "АдресТерминалаПосетителя" для терминалов "N_0"
            gEmplAddrTerm(0) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "EmplAddrTerm(1)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "АдресТерминалаПосетителя" для терминалов "N_1"
            gEmplAddrTerm(1) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "EmplAddrTerm(2)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "АдресТерминалаПосетителя" для терминалов "N_2"
            gEmplAddrTerm(2) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "EmplAddrTerm(3)" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию "АдресТерминалаПосетителя" для терминалов "N_3"
            gEmplAddrTerm(3) = frmTableSystem.grdTableSystem.Text
        
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "DefaultTime" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить умалчиваемый временной интервал при АвтоРегистрации Служащих
            gDefaultTime = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "DefaultStatus" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить умалчиваемое значение ячейки "Status"
            '  в "Таблице персон" при АвтоРегистрации Служащих
            gDefaultStatus = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "DefaultCalendar" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить умалчиваемое значение ячейки "Calendar"
            '  в "Таблице персон" при АвтоРегистрации Служащих
            gDefaultCalendar = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "NewCalend<='/*'" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить признак необходимости автоматического формирования
            '  опций выходных дней в "Таблице календаря" для Нового Года
            gHolidays = Trim(frmTableSystem.grdTableSystem.Text)
        
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "PrintSIAName" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Имя Компании-Владельца для
            '   его печати на Чековой Ленте
            gPrintSIAName = Trim(frmTableSystem.grdTableSystem.Text)
        
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "TalonLength" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Длину въездного/входного талона
            '   до линии отреза/отрыва
            gTalonLength = Trim(frmTableSystem.grdTableSystem.Text)
            If gTalonLength < 0 Then gTalonLength = 0
        
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "CreatePersonCode" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Опцию  автоматического формирования полей
            '  "PersonCode" и "Info"
            gCreatePersonCode = Trim(frmTableSystem.grdTableSystem.Text)
      
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "MinCount" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Минимальный возможный номер при автоматическом
            '  формировании поля "PersonCode"
            gMinCount = Trim(frmTableSystem.grdTableSystem.Text)
            If gMinCount < 10 Then gMinCount = 10
            
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "MaxCount" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Максимальный возможный номер при автоматическом
            '  формировании поля "PersonCode"
            gMaxCount = Trim(frmTableSystem.grdTableSystem.Text)
            If gMaxCount > 99 Then gMaxCount = 99
            
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "DisplayDiscount" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Уменьшение показаний количества свободных мест
            '  на дисплее по сравнению со счетчиком свободных мест
            '  (для устранения конфликтов при нескольких
            '  входах/въездах)
            gDisplayDiscount = Trim(frmTableSystem.grdTableSystem.Text)
            If gDisplayDiscount < 0 Then gDisplayDiscount = 0
            
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "TimeShare" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить Опцию разделения времени (параллельное вып. процедур)
            gTimeShare = Trim(frmTableSystem.grdTableSystem.Text)

        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "BeepSound" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить длительность звукового сигнала
            gBeepSound = Trim(frmTableSystem.grdTableSystem.Text)
        
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "DownLoadMSBase" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
            frmTableSystem.grdTableSystem.Col = 1
            'Установить опцию автоматической перезаписи "Таблицы протокола" в Базу
            gMSBase = Trim(frmTableSystem.grdTableSystem.Text)
        End If
            'Текущий столбец "Системной таблицы" = 2 (Тип)
        frmTableSystem.grdTableSystem.Col = 2
            'Тип="03" - Preprocessor (Препроцессор)
        If Left(Trim(frmTableSystem.grdTableSystem.Text), 2) = "03" Then
            'Количество Процессоров в локальной сети (исключая собственный)
            gNetPreprocNum = gNetPreprocNum + 1
            'Переопределить размерность массива Имен Процессоров
            '  локальной сети (исключая собственный Препроцессор)
            ReDim Preserve gSocketNet(gNetPreprocNum) As String
            If gPreprocIndex = intRowNumSys Then
            'Имя Процессора локальной сети (исключая собственный Препроцессор)
                gSocketNet(gNetPreprocNum) = gHost
            'Текущий столбец "Системной таблицы"=1(Константа)
                frmTableSystem.grdTableSystem.Col = 1
            'Действительное имя Процессора локальной сети - замена Псевдонима
                gPreprocName = Trim(frmTableSystem.grdTableSystem.Text)
            Else
            'Текущий столбец "Системной таблицы"=1(Константа)
                frmTableSystem.grdTableSystem.Col = 1
            'Имя Процессора локальной сети (исключая собственный Препроцессор)
                gSocketNet(gNetPreprocNum) = Trim(frmTableSystem.grdTableSystem.Text)
            End If
        End If
    Next

            ' Если это Препроцессор
    If gPreprocName <> "" Then
            'Копирование "Таблицы персон" из "Host Computer'a"
        Call TablePersonCopy
    End If
            ' Загрузить не показывая форму "frmTablePerson"
    Load frmTablePerson
            ' Если это "Host Computer"
    If gPreprocName = "" Then
            ' Получить ссылку на существующий в "Host Computer'e"
            '   объект ActiveX.EXE
        Set objTablePerson = New XTablePerson
            ' Получить ссылку на интерфейсы, объявленные для
            '   объекта "FlexGrid" ("Таблица Персон")
        Set gTablePerson = objTablePerson
            
            ' Если имеются Препроцессоры в локальной сети
        If gNetPreprocNum > 0 Then
            ' Создание объекта MSMQQueueInfo для управления
            '  очередью ПЕРЕДАВАЕМЫХ СООБЩЕНИЙ
            Set qInfoOutput = New MSMQQueueInfo
            ' Создать экземпляр объекта ОТСЫЛАЕМОЕ СООБЩЕНИЕ
            Set qMsgOutput = New MSMQMessage
            ' Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
            qMsgOutput.Label = gHost
            qMsgOutput.Body = "Host Started"
        
            'По всем элементам массива Имен Процессоров локальной сети
            For intCount = 1 To gNetPreprocNum
            ' Установить путь к очереди ПЕРЕДАВАЕМЫХ СООБЩЕНИЙ
                qInfoOutput.FormatName = "DIRECT=OS:" + _
                gSocketNet(intCount) + "\Private$\GeneralQueue"
            ' Открыть очередь сообщений с параметрами (для передачи
            '   сообщений, доступ к очереди разрешен всем)
                Set qQueueOutput = qInfoOutput.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            ' Отослать СООБЩЕНИЕ
                qMsgOutput.Send qQueueOutput
            ' Закрыть очередь СООБЩЕНИЙ
                qQueueOutput.Close
            Next
            
        End If
        
            ' Если это Препроцессор
    Else
            'Признак необходимости сжатия "Таблицы персон" не установлен:
            '   "Препроцессор" использует "Таблицу персон" "Host Computer'а"
        If gCompresTablPers = 0 Then
            ' Запустить удаленное приложение в "Host Computer'e",
            '   если оно еще не выполняется, и получить ссылку на
            '   объект ActiveX.EXE в нем
            Set objTablePerson = CreateObject("Sel_2Server.XTablePerson")
            ' Получить ссылку на интерфейсы, объявленные для
            '   объекта "FlexGrid" ("Таблица Персон")
            Set gTablePerson = objTablePerson
        
            'Признак необходимости сжатия "Таблицы персон" установлен:
            '   "Препроцессор" использует собственную "Таблицу персон"
        Else
            ' Получить ссылку на существующий в "Host Computer'e"
            '   объект ActiveX.EXE
            Set objTablePerson = New XTablePerson
            ' Получить ссылку на интерфейсы, объявленные для
            '   объекта "FlexGrid" ("Таблица Персон")
            Set gTablePerson = objTablePerson
        End If
'ОТЛАДКА
'Set objTablePerson = New XTablePerson
'Set gTablePerson = objTablePerson
        
            ' Создание объекта MSMQQueueInfo для управления
            '  очередью ПЕРЕДАВАЕМЫХ СООБЩЕНИЙ
        Set qInfoOutput = New MSMQQueueInfo
            ' Создать экземпляр объекта ОТСЫЛАЕМОЕ СООБЩЕНИЕ
        Set qMsgOutput = New MSMQMessage
            ' Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
        qMsgOutput.Label = gPreprocName
            ' Установить путь к очереди ПЕРЕДАВАЕМЫХ СООБЩЕНИЙ
        qInfoOutput.FormatName = "DIRECT=OS:" + gHost + "\Private$\GeneralQueue"
            ' Открыть очередь сообщений с параметрами (для передачи
            '   сообщений, доступ к очереди разрешен всем)
        Set qQueueOutput = qInfoOutput.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            ' Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
        qMsgOutput.Body = "Time"
            ' Отослать СООБЩЕНИЕ
        qMsgOutput.Send qQueueOutput
            ' Если имеется дисплей-указатель количества свободных мест
            '   на Автостоянке
        If gParkingPlaceNum <> 0 Then
            ' Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
            qMsgOutput.Body = "ParkFreePlaces "
            ' Отослать СООБЩЕНИЕ
            qMsgOutput.Send qQueueOutput
        End If
            ' Если имеется дисплей-указатель количества свободных мест
            '   на Предприятии
        If gAccessPlaceNum <> 0 Then
            ' Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
            qMsgOutput.Body = "AcceFreePlaces "
            ' Отослать СООБЩЕНИЕ
            qMsgOutput.Send qQueueOutput
        End If
            ' Закрыть очередь СООБЩЕНИЙ
        qQueueOutput.Close
        
            ' Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
        qMsgOutput.Body = "Preprocessor Started"
            'По всем элементам массива Имен Процессоров локальной сети
        For intCount = 1 To gNetPreprocNum
            ' Установить путь к очереди ПЕРЕДАВАЕМЫХ СООБЩЕНИЙ
            qInfoOutput.FormatName = "DIRECT=OS:" + _
            gSocketNet(intCount) + "\Private$\GeneralQueue"
            ' Открыть очередь сообщений с параметрами (для передачи
            '   сообщений, доступ к очереди разрешен всем)
            Set qQueueOutput = qInfoOutput.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            ' Отослать СООБЩЕНИЕ
            qMsgOutput.Send qQueueOutput
            ' Закрыть очередь СООБЩЕНИЙ
            qQueueOutput.Close
        Next
        
    End If
            ' Загрузить не показывая форму "frmTableCalendar"
    Load frmTableCalendar
            ' Загрузить не показывая форму "frmTableTime"
    Load frmTableTime
            ' Загрузить не показывая форму "frmTableTerminal"
    Load frmTableTerminal
             ' Если выбрана Опция Автоматического управления терминалами
             '  и установлен признак "Автостоянка", то Загрузить не показывая формы
             '  "frmDataParkingIn", "frmDataParkingOut", "frmDataParkingInfo,
             '  "frmDataParkingServ и "frmMinus"
    If (frmDemo.chkTerm(0).Value = 1 And intParking(0) = 1 Or _
    frmDemo.chkTerm(1).Value = 1 And intParking(1) = 1 Or _
    frmDemo.chkTerm(2).Value = 1 And intParking(2) = 1 Or _
    frmDemo.chkTerm(3).Value = 1 And intParking(3) = 1) And _
    frmDemo.optAutomatic = True Then
        Load frmDataParkingIn
        Load frmDataParkingOut
            'Изменение надписи на кнопке разрешения выезда для Специальных
            '  Клиентов с константной оплатой
        frmDataParkingOut.cmdOutConst.Caption = "Sant=""" + Str(gTariffConst) + """"
        Load frmDataParkingInfo
        Load frmDataParkingServ
        Load frmMinus
            'Сделать доступным меню "Parking" формы "frmDemo"
        frmDemo.mnuParking.Enabled = True
           'Сделать видимыми элементы управления "Автостоянкой"
        frmDemo.imgParkingIn.Visible = True
        frmDemo.imgParkingOut.Visible = True
        frmDemo.imgParkingInfo.Visible = True
        frmDemo.imgParkingServ.Visible = True
            'Опция "Печать Документа" при Регистрации/Исключении Клиента Автостоянки
        frmDataParkingIn.chkDocument.Value = intDocument
        frmDataParkingOut.chkDocument.Value = intDocument
        frmDataParkingInfo.chkDocument.Value = intDocument
        frmDataParkingServ.chkDocument.Value = intDocument
            'Открыть последовательный порт для ПРОСТОГО ЧЕКОВОГО ПРИНТЕРА
        If gDocument = 1 Or gDocument = 3 Or gDocument = 5 Or gDocument = 7 Then
            frmDemo.prtPortDocument.PortOpen = True
        End If
            'Открыть последовательный порт для ПРИНТЕРА ШТРИХ-КОДА
        If gDocument = 2 Or gDocument = 3 Or gDocument = 6 Or gDocument = 7 Then
            frmDemo.prtPortBarCode.PortOpen = True
        End If
            'Открыть последовательный порт для дисплея-указателя свободных мест
        If gParkingPlaceNum <> 0 Then
            frmDemo.prtPortDisplay.PortOpen = True
        End If
            'Установить Пароль операторов
        frmDataParkingIn.txtParole.Tag = strParole
        frmDataParkingOut.txtParole.Tag = strParole
        frmDataParkingInfo.txtParole.Tag = strParole
            'Установить время "TimeOut' для ввода пароля
        frmDataParkingIn.tmrParoleTimeOut.Interval = frmDemo.tmrPasswTimeOut.Interval
        frmDataParkingOut.tmrParoleTimeOut.Interval = frmDemo.tmrPasswTimeOut.Interval
        frmDataParkingInfo.tmrParoleTimeOut.Interval = frmDemo.tmrPasswTimeOut.Interval
           'Сделать видимыми элементы управления "Автостоянкой"
        If frmDemo.chkTerm(0).Value = 1 And intParking(0) = 1 Then
            If gTermInp = 0 Then frmDemo.imgParkingInData(0).Visible = True
            If gTermOut = 0 Then frmDemo.imgParkingOutData(0).Visible = True
            frmDemo.imgParkingInfoData(0).Visible = True
        End If
        If frmDemo.chkTerm(1).Value = 1 And intParking(1) = 1 Then
            If gTermInp = 1 Then frmDemo.imgParkingInData(1).Visible = True
            If gTermOut = 1 Then frmDemo.imgParkingOutData(1).Visible = True
            frmDemo.imgParkingInfoData(1).Visible = True
        End If
        If frmDemo.chkTerm(2).Value = 1 And intParking(2) = 1 Then
            If gTermInp = 2 Then frmDemo.imgParkingInData(2).Visible = True
            If gTermOut = 2 Then frmDemo.imgParkingOutData(2).Visible = True
            frmDemo.imgParkingInfoData(2).Visible = True
        End If
        If frmDemo.chkTerm(3).Value = 1 And intParking(3) = 1 Then
            If gTermInp = 3 Then frmDemo.imgParkingInData(3).Visible = True
            If gTermOut = 3 Then frmDemo.imgParkingOutData(3).Visible = True
            frmDemo.imgParkingInfoData(3).Visible = True
        End If
    Else
                'Сделать недоступным меню "Parking" формы "frmDemo"
        frmDemo.mnuParking.Enabled = False
           'Сделать невидимыми элементы управления "Автостоянкой"
        frmDemo.imgParkingIn.Visible = False
        frmDemo.imgParkingOut.Visible = False
        frmDemo.imgParkingInfo.Visible = False
        frmDemo.imgParkingServ.Visible = False
    End If
           
             ' Если выбрана Опция Автоматического управления терминалами
             '  и установлен признак "Посетитель", то Загрузить не показывая формы
             '  "frmDataAccessIn", "frmDataAccessOut", "frmDataAccessInfo",
             '  "frmDataAccessServ", "frmLease" и "frmMinus"
    If (frmDemo.chkTerm(0).Value = 1 And intAccess(0) = 1 Or _
    frmDemo.chkTerm(1).Value = 1 And intAccess(1) = 1 Or _
    frmDemo.chkTerm(2).Value = 1 And intAccess(2) = 1 Or _
    frmDemo.chkTerm(3).Value = 1 And intAccess(3) = 1) And _
    frmDemo.optAutomatic = True Then
        Load frmDataAccessIn
        Load frmDataAccessOut
            'Изменение надписи на кнопке разрешения выезда для Специальных
            '  Клиентов с константной оплатой
        frmDataAccessOut.cmdOutConst.Caption = "Sant=""" + Str(gTariffConst) + """"
        Load frmDataAccessInfo
        Load frmDataAccessServ
        Load frmLease
        Load frmMinus
            'Сделать доступным меню "Access" формы "frmDemo"
        frmDemo.mnuAccess.Enabled = True
           'Сделать видимыми элементы управления "Предприятием"
        frmDemo.imgAccessIn.Visible = True
        frmDemo.imgAccessOut.Visible = True
        frmDemo.imgAccessInfo.Visible = True
        frmDemo.imgAccessServ.Visible = True
            'Опция "Печать Документа" при Регистрации/Исключении Клиента
        frmDataAccessIn.chkDocument.Value = intDocument
        frmDataAccessOut.chkDocument.Value = intDocument
        frmDataAccessInfo.chkDocument.Value = intDocument
        frmDataAccessServ.chkDocument.Value = intDocument
            'Открыть последовательный порт для ПРОСТОГО ЧЕКОВОГО ПРИНТЕРА
        If (gDocument = 1 Or gDocument = 3 Or gDocument = 5 Or gDocument = 7) And _
        frmDemo.prtPortDocument.PortOpen = False Then
            frmDemo.prtPortDocument.PortOpen = True
        End If
            'Открыть последовательный порт для ПРИНТЕРА ШТРИХ-КОДА
        If (gDocument = 2 Or gDocument = 3 Or gDocument = 6 Or gDocument = 7) And _
        frmDemo.prtPortBarCode.PortOpen = False Then
            frmDemo.prtPortBarCode.PortOpen = True
        End If
            'Открыть последовательный порт для дисплея-указателя свободных мест
        If gAccessPlaceNum <> 0 And frmDemo.prtPortDisplay.PortOpen = False Then
            frmDemo.prtPortDisplay.PortOpen = True
        End If
            'Установить Пароль операторов
        frmDataAccessIn.txtParole.Tag = strParole
        frmDataAccessOut.txtParole.Tag = strParole
        frmDataAccessInfo.txtParole.Tag = strParole
            'Установить время "TimeOut' для ввода пароля
        frmDataAccessIn.tmrParoleTimeOut.Interval = frmDemo.tmrPasswTimeOut.Interval
        frmDataAccessOut.tmrParoleTimeOut.Interval = frmDemo.tmrPasswTimeOut.Interval
        frmDataAccessInfo.tmrParoleTimeOut.Interval = frmDemo.tmrPasswTimeOut.Interval
           'Сделать видимыми элементы управления "Предприятием"
        If frmDemo.chkTerm(0).Value = 1 And intAccess(0) = 1 Then
            If gTermInp = 0 Then frmDemo.imgAccessInData(0).Visible = True
            If gTermOut = 0 Then frmDemo.imgAccessOutData(0).Visible = True
            frmDemo.imgAccessInfoData(0).Visible = True
        ElseIf frmDemo.chkTerm(1).Value = 1 And intAccess(1) = 1 Then
            If gTermInp = 1 Then frmDemo.imgAccessInData(1).Visible = True
            If gTermOut = 1 Then frmDemo.imgAccessOutData(1).Visible = True
            frmDemo.imgAccessInfoData(1).Visible = True
        ElseIf frmDemo.chkTerm(2).Value = 1 And intAccess(2) = 1 Then
            If gTermInp = 2 Then frmDemo.imgAccessInData(2).Visible = True
            If gTermOut = 2 Then frmDemo.imgAccessOutData(2).Visible = True
            frmDemo.imgAccessInfoData(2).Visible = True
        ElseIf frmDemo.chkTerm(3).Value = 1 And intAccess(3) = 1 Then
            If gTermInp = 3 Then frmDemo.imgAccessInData(3).Visible = True
            If gTermOut = 3 Then frmDemo.imgAccessOutData(3).Visible = True
            frmDemo.imgAccessInfoData(3).Visible = True
        End If
    Else
                'Сделать недоступным меню "Access" формы "frmDemo"
        frmDemo.mnuAccess.Enabled = False
           'Сделать невидимыми элементы управления "Предприятием"
        frmDemo.imgAccessIn.Visible = False
        frmDemo.imgAccessOut.Visible = False
        frmDemo.imgAccessInfo.Visible = False
        frmDemo.imgAccessServ.Visible = False
    End If
           
             ' Если выбрана Опция Автоматического управления терминалами
             '  и установлен признак "Служащий", то Загрузить не показывая формы
             '  "frmDataEmployeIn", "frmDataEmployeOut" и "frmDataAccessInfo
    If (frmDemo.chkTerm(0).Value = 1 And intEmploye(0) = 1 Or _
    frmDemo.chkTerm(1).Value = 1 And intEmploye(1) = 1 Or _
    frmDemo.chkTerm(2).Value = 1 And intEmploye(2) = 1 Or _
    frmDemo.chkTerm(3).Value = 1 And intEmploye(3) = 1) And _
    frmDemo.optAutomatic = True Then
        Load frmDataEmployeIn
        Load frmDataEmployeOut
        Load frmDataEmployeInfo
            'Сделать доступным меню "Employe" формы "frmDemo"
        frmDemo.mnuEmploye.Enabled = True
           'Сделать видимыми элементы управления "Служащими" (Регистрация и пр.)
        frmDemo.imgEmployeInData.Visible = True
        frmDemo.imgEmployeOutData.Visible = True
        frmDemo.imgEmployeInfoData.Visible = True
            'Установить Пароль операторов
        frmDataEmployeIn.txtParole.Tag = strParole
        frmDataEmployeOut.txtParole.Tag = strParole
        frmDataEmployeInfo.txtParole.Tag = strParole
            'Установить время "TimeOut' для ввода пароля
        frmDataEmployeIn.tmrParoleTimeOut.Interval = frmDemo.tmrPasswTimeOut.Interval
        frmDataEmployeOut.tmrParoleTimeOut.Interval = frmDemo.tmrPasswTimeOut.Interval
        frmDataEmployeInfo.tmrParoleTimeOut.Interval = frmDemo.tmrPasswTimeOut.Interval
    Else
                'Сделать недоступным меню "Employe" формы "frmDemo"
        frmDemo.mnuEmploye.Enabled = False
    End If
           
           'Сделать видимой форму 'frmDemo"
    frmDemo.Visible = True
            'Установить фокус на опции "Dummy"
    If frmDemo.Visible = True Then frmDemo.chkDummy.SetFocus
    
            'Вычислить и запомнить текущую дату
    frmTableCalendar.Tag = Trim(Format(Now, "dd/mm/yyyy"))
            'Текущее время
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Часы
    intHour = Hour(gProtocol.strProtocTime)
    If intHour < 10 Then
        strHour = "0" + Trim(Str(intHour))
    Else
        strHour = Trim(Str(intHour))
    End If
            'Минуты
    intMinute = Minute(gProtocol.strProtocTime)
    If intMinute < 10 Then
        strMinute = "0" + Trim(Str(intMinute))
    Else
        strMinute = Trim(Str(intMinute))
    End If
            
            'Вывод Даты и Времени
    frmDemo.lblTime.Caption = "   " + frmTableCalendar.Tag + "   " _
    + strHour + ":" + strMinute
            
    
            'Запустить таймер минутного отсчета
    frmTableCalendar.tmrMinute.Enabled = True
            'Запрет опроса терминалов
    gTermContr = 0
            'Включить таймер циклического опроса "Controller'ов"
            '  командой ЧТЕНИЕ ПЕРСОНАЛЬНОГО КОДА
    If frmDemo.chkTerm(0).Value = 1 Or frmDemo.chkTerm(1).Value = 1 Or _
    frmDemo.chkTerm(2).Value = 1 Or frmDemo.chkTerm(3).Value = 1 Then
        frmDemo.tmrTermContr.Enabled = True
    End If
            'Вывод изображений терминалов в исходном (закрытом) состоянии
            '  и Открыть последовательные порты для отмеченных терминалов
    If frmDemo.chkTerm(0).Value = 1 Then
        frmDemo.prtPortC(0).PortOpen = True
        frmDemo.imgViewClose(0).Visible = True
    End If
    If frmDemo.chkTerm(1).Value = 1 Then
        frmDemo.prtPortC(1).PortOpen = True
        frmDemo.imgViewClose(1).Visible = True
    End If
    If frmDemo.chkTerm(2).Value = 1 Then
        frmDemo.prtPortC(2).PortOpen = True
        frmDemo.imgViewClose(2).Visible = True
    End If
    If frmDemo.chkTerm(3).Value = 1 Then
        frmDemo.prtPortC(3).PortOpen = True
        frmDemo.imgViewClose(3).Visible = True
    End If
            'Установить адреса "Controller'ов" при "Ручном"
            '  управления терминалами от электонной "Кнопки"
    If frmDemo.optManual = True Then
        frmDemo.cmdOpen(0).Tag = CByte(CInt(Trim(gAddrManual(0))))
        frmDemo.cmdOpen(0).Caption = Trim(gAddrManual(0))
        frmDemo.cmdOpen(1).Tag = CByte(CInt(Trim(gAddrManual(1))))
        frmDemo.cmdOpen(1).Caption = Trim(gAddrManual(1))
        frmDemo.cmdOpen(2).Tag = CByte(CInt(Trim(gAddrManual(2))))
        frmDemo.cmdOpen(2).Caption = Trim(gAddrManual(2))
        frmDemo.cmdOpen(3).Tag = CByte(CInt(Trim(gAddrManual(3))))
        frmDemo.cmdOpen(3).Caption = Trim(gAddrManual(3))
    End If
            
            ' Если имеется дисплей-указатель количества свободных мест
            '   на Автостоянке
    If gParkingPlaceNum <> 0 Then
            'Инициализация дисплея-указателя
        strMessage = "ParkFreePlaces=" + CStr(gParkFreePlaces)
        Call frmDemo.Display(strMessage)
            ' Если имеется дисплей-указатель количества свободных мест
            '   на Предприятии
    ElseIf gAccessPlaceNum <> 0 Then
            'Инициализация дисплея-указателя
        strMessage = "AcceFreePlaces=" + CStr(gParkFreePlaces)
        Call frmDemo.Display(strMessage)
    End If
            
            'Скрыть изображения терминалов в открытом состоянии
    frmDemo.imgViewOpen(0).Visible = False
    frmDemo.imgViewOpen(1).Visible = False
    frmDemo.imgViewOpen(2).Visible = False
    frmDemo.imgViewOpen(3).Visible = False
            ' Вывод "пустых" фотоизображений
    frmDemo.imgPhoto(0).Picture = LoadPicture("")
    frmDemo.imgPhoto(1).Picture = LoadPicture("")
    frmDemo.imgPhoto(2).Picture = LoadPicture("")
    frmDemo.imgPhoto(3).Picture = LoadPicture("")
            'Сбросить признак "Белый ключ"
    intWhite = 0
           'Начальный порядковый элемент в массивах управляющих элементов форм
    intControlIndex = 0
    
               'Обнулить номер текущей строки "Таблицы персон"
    frmDemo.tmrTermContr.Tag = 0
    frmDemo.lblInform(0).Tag = 0
    frmDemo.lblInform(1).Tag = 0
    frmDemo.lblInform(2).Tag = 0
    frmDemo.lblInform(3).Tag = 0
               'Разрешение опроса терминалов
    gTermContr = 1
            ' Цикл опроса "Controller'ов" терминалов
            '    (завершается при нажатии кнопки "cmdExit")
    Do While DoEvents()
            
            'Если объект ActiveX.EXE в "Host Computer'e" разрушен
        If objTablePerson Is Nothing Then
            ' Если это "Host Computer"
            If gPreprocName = "" Then
            'Восстановить объект в "Host Computer'e"
                Set objTablePerson = New XTablePerson
                Set gTablePerson = objTablePerson
            End If
        End If
               
               'Обнулить номер текущей строки "Таблицы персон"
        frmDemo.tmrTermContr.Tag = intControlIndex
        frmDemo.lblInform(intControlIndex).Tag = 0
            
            'Ждать от "Controller'ов" открытых портов
            '  получения ПЕРСОНАЛЬНОГО КОДА
            
        If frmDemo.prtPortC(intControlIndex).InBufferCount >= _
        gPersonCode(intControlIndex) Then
           'Полученные данные в приемный буфер для дальнейшей обработки
            vntBufferInput(intControlIndex) = frmDemo.prtPortC(intControlIndex).Input
            'Вызов "Начальной последовательности открытия терминала"
            Call InitialOpenTerminal(intControlIndex)
        End If
        
'''        DoEvents
        
            'Имеются Запросы на обслуживание терминалов
        If frmDemo.prtPortC(intControlIndex).Tag > 0 Then
            'Номер конечного элемента для порта, предшествующего
            ' текущему порту в массиве "Таблицы терминалов"
            intWork = (frmDemo.prtPortC(intControlIndex).CommPort - 2) * 15
            'По всем столбцам строки массива для текущего порта
            For intCount = 1 To 15 Step 1
            'Номер текущего элемента массива "Таблицы терминалов",
            '  хранящего Запрос
                intRequest = intWork + intCount
            'Запрос на "Начальную последовательность закрытия терминала"
                If Mid(gAddrPort(0, intRequest), 4) = "A" Then
            'Вызов "Начальной последовательности закрытия терминала"
                    Call InitialCloseTerminal(intControlIndex, intRequest)
            'Запрос на "Последовательность ожидания закрытия терминала"
                ElseIf Mid(gAddrPort(0, intRequest), 4) = "V" Then
            'Вызов "Последовательности ожидания закрытия терминала"
                    Call WaitCloseTerminal(intControlIndex, intRequest)
            'Запрос на "Последовательность открытия терминала"
            '  от электронной "Кнопки"
                ElseIf Mid(gAddrPort(0, intRequest), 4) = "1" Then
            'Вызов "Последовательности открытия терминала"
            '  от электронной "Кнопки"
                    Call ButtonOpenTerminal(intControlIndex, intRequest)
                End If
            Next
        End If
            
            'Bывести изображение закрытого терминала
        If frmDemo.tmrButton(intControlIndex).Tag = 1 Then
            PictureTerminalClose intControlIndex
            'Сбросить признак "TimeOut" для электронной "Кнопки"
            frmDemo.tmrButton(intControlIndex).Tag = 0
        End If
        
        If intControlIndex < 3 Then
            intControlIndex = intControlIndex + 1
        Else
            intControlIndex = 0
        End If
            'Порт с текущим номером закрыт - перейдти к следующему порту
        Do While frmDemo.prtPortC(intControlIndex).PortOpen = False And _
        frmDemo.chkSetup.Value = 1
          'Порядковый номер элемента в массивах управляющих элементов форм
            If intControlIndex < 3 Then
                intControlIndex = intControlIndex + 1
            Else
                intControlIndex = 0
                Exit Do
            End If
        Loop
        
    Loop
    
End Sub
            
            'Начальная последовательность открытия терминала
Private Sub InitialOpenTerminal(intIndex As Integer)
            'Номер порта, через который принимается
            '  ПЕРСОНАЛЬНЫЙ КОД от "Controller'a"
Dim vntReadPortNum As Variant
            'Код возврата при анализе результата АвтоРегистрации
Dim intAutoRegistrCode As Integer
            'Код возврата при анализе результата АвтоУдаления
Dim intAutoDeleteCode As Integer
            'Код возврата функций коррекции ячейки "Reserve" в "Таблице персон"
            '  после въезда/выезда Клиента Автостоянки
Dim intParkingCode As Integer
            'Код возврата функций коррекции ячейки "Reserve" в "Таблице персон"
            '  после входа/выхода Посетителя Предприятия
Dim intAccessCode As Integer
            'Код возврата функций коррекции ячейки "Name" в "Таблице персон"
            '  после входа/выхода Служащего Предприятия
Dim intEmployeCode As Integer
            'Код возврата при анализе статуса доступа
Dim intStatusCode As Integer
            'Код возврата при анализе дня доступа
Dim intCalendarCode As Integer
            'Код возврата при анализе времени доступа
Dim intTimeCode As Integer
            'Код возврата при анализе терминала доступа
Dim intTerminalCode As Integer
            'Код возврата при ожидании КВИТАНЦИИ или КОДА СОСТОЯНИЯ
Dim intScriptCode As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Номер элемента в массиве "Таблицы терминалов"
Dim intRequest As Integer
            'Рабочее поле
Dim vntWork As Variant
            
            'Переменнные для преобразования битовых данных
            ' из буфера в шестнадцатиричное представление
Dim intCicle As Integer
Dim strBuffer As String
Dim intBuffer1 As Integer
Dim intBuffer2 As Integer
            
            'Номер порта, получившего ПЕРСОНАЛЬНЫЙ КОД от "Controller'a"
    vntReadPortNum = frmDemo.prtPortC(intIndex).CommPort
            'Определить адрес "Controller'a", приславшего ПЕРСОНАЛЬНЫЙ КОД
    vntAddr = CByte(Asc(Left(vntBufferInput(intIndex), 1))) And CByte(15)
            'Номер текущего элемента в массиве "Таблицы терминалов"
    intRequest = (vntReadPortNum - 2) * 15 + vntAddr
            
            'Анализ состояния соответствующего "Controller'a"
            
            ' "Controller" ЛОГИЧЕСКИ выключен из системы
    If Mid(gAddrPort(0, intRequest), 1, 2) = "00" Then GoTo WaitCycle
            ' "Controller" ЗАНЯТ обслуживанием терминала и нет ожидания
            '  Оранжевого индикатора на считывателе - выход из процедуры
    If Mid(gAddrPort(0, intRequest), 4) <> "0" And _
    Mid(gAddrPort(0, intRequest), 4) <> "#" Then GoTo WaitCycle
            'Сбросить признак "Белый ключ"
    intWhite = 0
    
            'Выделить биты индикации считывателя из
            '  полученного ПЕРСОНАЛЬНОГО КОДА
    vntWork = CByte(Asc(Left(vntBufferInput(intIndex), 1))) And CByte(96)
            ' "Controller" не хранит PIN код (не Красный индикатор на считывателе)
    If vntWork <> 32 Then
            ' "Controller" обрабатывает ПЕРСОНАЛЬНЫЙ КОД, который хранится
            '   в его локальной памяти (Зеленый индикатор на считывателе) и нет
            '   ожидания Оранжевого индикатора на считывателе
        If vntWork = 0 And Mid(gAddrPort(0, intRequest), 4) <> "#" Then
            'Установить признак "Белый ключ"
            intWhite = 1
            GoTo Continue
            ' "Controller" выполняет команду ОТКРЫТЬ ТЕРМИНАЛ (от компьютера
            '  или свою собственную - при поступлении ПЕРСОНАЛЬНОГО КОДА,
            '  который хранится в локальной памяти "Controller'a") и ожидает
            '  Оранжевого индикатора на считывателе
        ElseIf vntWork = 0 And Mid(gAddrPort(0, intRequest), 4) = "#" Then
            GoTo WaitCycle
            ' "Controller" готов принять PIN код (Оранжевый индикатор на считывателе)
            '   и ожидает Оранжевого индикатора на считывателе
        ElseIf vntWork = 64 And Mid(gAddrPort(0, intRequest), 4) = "#" Then
            'Сбросить установленный признак ЗАНЯТОГО "Controller'a", у
            '   которого ожидается Оранжевый индикатор на считывателе
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "0"
            GoTo WaitCycle
            ' "Controller" готов принять PIN код (Оранжевый индикатор на считывателе)
            '   и не ожидает Оранжевого индикатора на считывателе
        ElseIf vntWork = 64 And Mid(gAddrPort(0, intRequest), 4) = "0" Then
            GoTo WaitCycle
            ' "Controller" находится в состоянии ПРОГРАММИРОВАНИЕ
            '  (горят Красный и Оранжевый индикаторы на считывателе)
        ElseIf vntWork = 96 Then
            'Сбросить возможный признак ЗАНЯТОГО "Controller'a", у
            '   которого ожидается Оранжевый индикатор на считывателе
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "0"
            GoTo WaitCycle
        End If
            ' "Controller" хранит PIN код (Красный индикатор на считывателе)
            '   и ожидает Оранжевого индикатора на считывателе
    ElseIf vntWork = 32 And Mid(gAddrPort(0, intRequest), 4) = "#" Then
            'Сбросить установленный признак ЗАНЯТОГО "Controller'a", у
            '   которого ожидается Оранжевый индикатор на считывателе
        gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "0"
        GoTo Continue
            ' "Controller" хранит PIN код (Красный индикатор на считывателе)
            '   и не ожидает Оранжевого индикатора на считывателе
    ElseIf vntWork = 32 And Mid(gAddrPort(0, intRequest), 4) = "0" Then
        GoTo Continue
    End If
    
Continue:
            
            'Унификация структуры ПЕРСОНАЛЬНОГО КОДА
            
            'У "Controller'f" было включено питание
            '  или установлен запрещенный нулевой адрес
    If vntAddr = 0 Then
            'Нулевые ПЕРСОНАЛЬНЫЙ КОД и адрес "Controller'a"
        vntBufferInput(intIndex) = "0000000000000000"
            'ПЕРСОНАЛЬНЫЙ КОД в надпись метки "N_?"
        frmDemo.lblInform(intIndex).Caption = vntBufferInput(intIndex)
            'Метка "N_?" - (белый фон)
        frmDemo.lblInform(intIndex).BackColor = vbWhite
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
        frmDemo.BeepSound
            'ПЕРСОНАЛЬНЫЙ КОД и адрес "Controller'a"
        gProtocol.strProtocName = vntBufferInput(intIndex)
            'ПЕРСОНАЛЬНЫЙ КОД и адрес "Controller'a"
        gProtocol.strProtocPersonCode = vntBufferInput(intIndex)
            'Статус
        gProtocol.strProtocStatus = ""
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
        gProtocol.strProtocReserve = "POWER ON or ADDR=0"
            'Записать строку в файл "Таблицы протокола"
        GoTo Protocol
            'К "Controller'y" подсоединен считыватель "PROXIMITY GP30"
    ElseIf CByte(Asc(Right(vntBufferInput(intIndex), 1))) = CByte(3) Then
        vntBufferInput(intIndex) = Left(vntBufferInput(intIndex), 1) + "00000" + _
        Mid(vntBufferInput(intIndex), 4, 10)
            'К "Controller'y" подсоединен считыватель штрих-кода"VS800"
    ElseIf Right(vntBufferInput(intIndex), 1) = "*" Then
        vntBufferInput(intIndex) = Left(vntBufferInput(intIndex), 1) + "00000" + _
        Mid(vntBufferInput(intIndex), 5, 10)
            'К "Controller'y" подсоединен считыватель "TM DALLAS" для "1990A"
    ElseIf CByte(Asc(Mid(vntBufferInput(intIndex), 9, 1))) = CByte(1) Then
            'Преобразовать данные из буфера терминала в шестнадцатиричный вид
        strBuffer = ""
        intCicle = 1
        Do While intCicle <= 6
            intBuffer1 = (CByte(Asc(Mid(vntBufferInput(intIndex), intCicle + 9, 1))) And CByte(240)) / 16
            intBuffer2 = CByte(Asc(Mid(vntBufferInput(intIndex), intCicle + 9, 1))) And CByte(15)
            strBuffer = Hex(intBuffer1) + Hex(intBuffer2) + strBuffer
            intCicle = intCicle + 1
        Loop
        vntBufferInput(intIndex) = Left(vntBufferInput(intIndex), 1) + "000" + Trim(strBuffer)
            'К "Controller'у" подсоединен недопустимый тип считывателя
    ElseIf vntAddr <> 0 Then
            'ПРЕДУПРЕЖДЕНИЕ и адрес "Controller'a" в надпись метки "N_?"
        frmDemo.lblInform(intIndex).Caption = CStr(vntAddr) + "||" + "UndefinedErr"
            'Метка "N_?" - (белый фон)
        frmDemo.lblInform(intIndex).BackColor = vbWhite
            'ПРЕДУПРЕЖДЕНИЕ и адрес "Controller'a"
        gProtocol.strProtocName = CStr(vntAddr) + "||" + "ErReaderType"
            'ПРЕДУПРЕЖДЕНИЕ и адрес "Controller'a"
        gProtocol.strProtocPersonCode = CStr(vntAddr) + "||" + "CoflictComm"
            'Статус
        gProtocol.strProtocStatus = ""
            'Статус
        gProtocol.strProtocStatus = ""
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
        gProtocol.strProtocReserve = "UNDEFINED ERROR"
            'Записать строку в файл "Таблицы протокола"
        GoTo Protocol
    End If
            
            'Метки сообщения с предупреждением - убрать с экрана
    frmDemo.lblErrorInpOut(intIndex).Visible = False
    frmDemo.lblErrorBarCodePrinter.Visible = False
            
            'Увеличить текущее значение атрибута
            '  доступности "Таблицы персон"
    gTablePerson.AccessPlus
            
            'Выделить номер считывателя из полученного ПЕРСОНАЛЬНОГО КОДА
    vntWork = CByte(Asc(Left(vntBufferInput(intIndex), 1))) And CByte(16)
    If vntWork = 16 Then vntWork = 1
            'ПЕРСОНАЛЬНЫЙ КОД поступил от "Controller'a" и "Port'a"
            '  АвтоРегистрации/Удаления
    If vntAddr * 10 + vntReadPortNum = intAutoRegDel(intIndex) Then
            'ПЕРСОНАЛЬНЫЙ КОД без адреса "Controller'a"
        gProtocol.strProtocPersonCode = "0" + Mid(vntBufferInput(intIndex), 2, _
        gPersonCode(intIndex) - 1)
            'Номер считывателя="0" - АвтоРегистрация
        If vntWork = 0 Then
            'Вызов процедуры-функции инициализации ввода ПЕРСОНАЛЬНОГО КОДА
            '  и других данных при АвтоРегистрации Клиента на Автостоянке
            If intParking(intIndex) = 1 Then
                intAutoRegistrCode = _
                frmDemo.AutoParkReg(gProtocol.strProtocPersonCode, intIndex)
                GoTo WaitCycle
            'Вызов процедуры-функции инициализации ввода ПЕРСОНАЛЬНОГО КОДА
            '  и других данных при АвтоРегистрации Посетителя Предприятия
            ElseIf intAccess(intIndex) = 1 Then
                intAutoRegistrCode = _
                frmDemo.AutoAcceReg(gProtocol.strProtocPersonCode, intIndex)
                GoTo WaitCycle
            'Вызов процедуры-функции инициализации ввода ПЕРСОНАЛЬНОГО КОДА
            '  и других данных при АвтоРегистрации Служащего Предприятия
            ElseIf intEmploye(intIndex) = 1 Then
                intAutoRegistrCode = _
                frmDemo.AutoEmplReg(gProtocol.strProtocPersonCode)
                GoTo WaitCycle
            End If
            'Номер считывателя="1" - АвтоУдаление
        Else
            'Вызов процедуры-функции инициализации ввода ПЕРСОНАЛЬНОГО КОДА
            '  и других данных при АвтоУдалении Клиента с Автостоянки
            If intParking(intIndex) = 1 Then
                intAutoDeleteCode = _
                frmDemo.AutoParkDel(gProtocol.strProtocPersonCode, intIndex)
                GoTo WaitCycle
            'Вызов процедуры-функции инициализации ввода ПЕРСОНАЛЬНОГО КОДА
            '  и других данных при АвтоУдалении Посетителя Предприятия
            ElseIf intAccess(intIndex) = 1 Then
                intAutoDeleteCode = _
                frmDemo.AutoAcceDel(gProtocol.strProtocPersonCode, intIndex)
                GoTo WaitCycle
            'Вызов процедуры-функции инициализации ввода ПЕРСОНАЛЬНОГО КОДА
            '  и других данных при АвтоУдалении Служащего Предприятия
            ElseIf intEmploye(intIndex) = 1 Then
                intAutoDeleteCode = _
                frmDemo.AutoEmplDel(gProtocol.strProtocPersonCode)
                GoTo WaitCycle
            End If
        End If
    End If
    
'''            'Установлена опция "Шлюз" для терминалов порта
'''    If gSluice(intIndex) <> 0 Then
'''            'Блокировать все "Controller'ы" порта - БЛОКИРОВКА с групповым адресом
'''        frmDemo.prtPortC(intIndex).Output = Chr(192)
'''             'Ждать завершения передачи команды БЛОКИРОВКА
'''        Do
'''        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
'''            'Сбросить соответствующий "Controller" - СБРОС с собственным адресом
'''        frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
'''             'Ждать завершения передачи команды СБРОС
'''        Do
'''        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
'''    End If

            'ПЕРСОНАЛЬНЫЙ КОД в надпись метки "N_?"
    frmDemo.lblInform(intIndex).Caption = CStr(vntAddr) + "||" + _
    Mid(vntBufferInput(intIndex), 7, gPersonCode(intIndex) - 6)
    
            'ПЕРСОНАЛЬНЫЙ КОД без адреса "Controller'a"
    gProtocol.strProtocPersonCode = "0" + Mid(vntBufferInput(intIndex), 2, _
    gPersonCode(intIndex) - 1)
        'Текущий столбец "Таблицы персон" = 1 (Персональный код)
    gTablePerson.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = intRowNum
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
        If Trim(gTablePerson.Text) = gProtocol.strProtocPersonCode Then
            'Досрочный выход из цикла
            Exit For
        End If
    Next
            
            ' Очистка поля фотоизображения
    If frmDemo.chkPhoto(intIndex).Value = 1 Then
        frmDemo.imgPhoto(intIndex).Picture = LoadPicture("")
    End If
            
            'Протоколирование события
            
            'Недействительный ПЕРСОНАЛЬНЫЙ КОД
    If intRowNum = gTablePerson.Rows Then
                'Резервированное Имя (когда получен неизвестный ПЕРСОНАЛЬНЫЙ КОД)
        gProtocol.strProtocName = "@"
            'Установлен признак "Белый ключ"
        If intWhite = 1 Then
            'Примечание - "Белый ключ"
            
            'Номер считывателя="0"
            If vntWork = 0 Then
                gProtocol.strProtocReserve = "0/" + CStr(vntAddr) + "/" + _
                CStr(vntReadPortNum) + "  WHITE KEY"
            'Номер считывателя="1"
            Else
                gProtocol.strProtocReserve = "1/" + CStr(vntAddr) + "/" + _
                CStr(vntReadPortNum) + "  WHITE KEY"
            End If
            
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
            frmDemo.BeepSound
            'Сбросить признак "Белый ключ"
            intWhite = 0
            'Признак "Белый ключ" не установлен - (Неверный ПЕРСОНАЛЬНЫЙ КОД)
        Else
            'Номер считывателя="0"
            If vntWork = 0 Then
                gProtocol.strProtocReserve = "0/" + CStr(vntAddr) + "/" + _
                CStr(vntReadPortNum) + "  INVALID KEY"
            'Номер считывателя="1"
            Else
                gProtocol.strProtocReserve = "1/" + CStr(vntAddr) + "/" + _
                CStr(vntReadPortNum) + "  INVALID KEY"
            End If
        
        End If
            'Статус
        gProtocol.strProtocStatus = ""
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Установить признак ЗАНЯТОГО "Controller'a", у которого
            '   ожидается Оранжевый индикатор на считывателе
        gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            'Метка "N_?" - (белый фон)
        frmDemo.lblInform(intIndex).BackColor = vbWhite
        GoTo ResetController
            'Действительный ПЕРСОНАЛЬНЫЙ КОД
    Else
            
            'Запомнить номер текущей строки "Таблицы персон"
        frmDemo.lblInform(intIndex).Tag = intRowNum
                
            'Текущий столбец "Таблицы персон" = 0 (Персона или Терминал)
        gTablePerson.Col = 0
        gProtocol.strProtocName = gTablePerson.Text
            'Текущий столбец "Таблицы персон" = 2 (Статус)
        gTablePerson.Col = 2
        gProtocol.strProtocStatus = Trim(gTablePerson.Text)
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
        intHour = Hour(gProtocol.strProtocTime)
        intMinute = Minute(gProtocol.strProtocTime)
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание из надписи метки "N_?"
        gProtocol.strProtocReserve = frmDemo.lblInform(intIndex).Caption
    
            'Вывод отмеченного фотоизображения из папки "Photo" изображений
        If frmDemo.chkPhoto(intIndex).Value = 1 Then
            'Перехват обработки ошибок ввода/вывода
            On Error GoTo PhotoError
            'Не установлен Признак "Автостоянка" или платное "Предприятие"
            If intParking(intIndex) = 0 And intAccess(intIndex) = 0 Then
                frmDemo.imgPhoto(intIndex).Picture = LoadPicture("C:\Photo\" + _
                Trim(Left(gProtocol.strProtocName, 15)) + ".bmp")
            'Установлен Признак "Автостоянка" или платное "Предприятие"
            Else
                frmDemo.imgPhoto(intIndex).Picture = LoadPicture("C:\Photo\" + _
                Trim(gProtocol.strProtocName) + ".bmp")
            End If
            GoTo PhotoOK
PhotoError:
            Resume PhotoOK
PhotoOK:
            On Error GoTo 0
        End If
            
            'Анализ статуса доступа
        intStatusCode = StatusCode()
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            'Установлен признак "Белый ключ", но доступ
            '  не через локальный терминал - некорректность
        If intWhite = 1 And intStatusCode <> 0 Then
            
            'Номер считывателя="0"
            If vntWork = 0 Then
                gProtocol.strProtocReserve = "0/" + CStr(vntAddr) + "/" + _
                CStr(vntReadPortNum) + "  WHITE KEY"
            'Номер считывателя="1"
            Else
                gProtocol.strProtocReserve = "1/" + CStr(vntAddr) + "/" + _
                CStr(vntReadPortNum) + "  WHITE KEY"
            End If
            
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
            frmDemo.BeepSound
            'Сбросить признак "Белый ключ"
            intWhite = 0
            'Установить признак ЗАНЯТОГО "Controller'a", у которого
            '   ожидается Оранжевый индикатор на считывателе
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            'Метка "N_?" - (белый фон)
            frmDemo.lblInform(intIndex).BackColor = vbWhite
            'Не Черный ключ
            If intStatusCode <> 2 Then GoTo ResetController
        
        End If
            'Доступ через локальный терминал - НЕ ДОПУСКАЕТСЯ для Автостоянки
            '  и платного Предприятия
        If intStatusCode = 0 And intParking(intIndex) = 0 And intAccess(intIndex) = 0 Then
            'Установить признак ЗАНЯТОГО "Controller'a", у которого
            '   ожидается Оранжевый индикатор на считывателе
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            'Метка "N_?" - (синий фон)
            frmDemo.lblInform(intIndex).BackColor = vbBlue
            GoTo ResetController
            'Черный ключ
        ElseIf intStatusCode = 2 Then
            'Протоколирование события - (Черный ключ)
            
            'Номер считывателя="0"
            If vntWork = 0 Then
                gProtocol.strProtocReserve = "0/" + CStr(vntAddr) + "/" + _
                CStr(vntReadPortNum) + "  BLACK KEY"
            'Номер считывателя="1"
            Else
                gProtocol.strProtocReserve = "1/" + CStr(vntAddr) + "/" + _
                CStr(vntReadPortNum) + "  BLACK KEY"
            End If
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
            frmDemo.BeepSound
            'Установить признак ЗАНЯТОГО "Controller'a", у которого
            '   ожидается Оранжевый индикатор на считывателе
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            'Метка "N_?" - (красный фон)
            frmDemo.lblInform(intIndex).BackColor = vbRed
            GoTo ResetController
            'Статус "Relay" - Доступ при ручном подтверждении охраны и "Автоматическом"
            '  управлении терминалами - НЕ ДОПУСКАЕТСЯ для Автостоянки, платного
            '  Предприятия или Проходной для Служащих
        ElseIf intStatusCode = 3 And frmDemo.optAutomatic = True And _
        intParking(intIndex) = 0 And intAccess(intIndex) = 0 And intEmploye(intIndex) = 0 Then
            'Текущий столбец "Таблицы персон" = 0 (Персона или Терминал)
            gTablePerson.Col = 0
            ' Электронная "Кнопка" не хранит адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала
            If Trim(gTablePerson.Text) <> "Dallas" And _
            frmDemo.cmdOpen(intIndex).Tag = 0 Then
'''            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
'''                frmDemo.BeepSound
            'Записать в электронную "Кнопку" адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала
                frmDemo.cmdOpen(intIndex).Tag = vntAddr
                frmDemo.cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            'Сделать электронную "Кнопку" временно доступной
                frmDemo.cmdOpen(intIndex).Enabled = True
            'Установить признак ЗАНЯТОГО "Controller'a", у которого
            '   ожидается Оранжевый индикатор на считывателе
                gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            'Метка "N_?" - (желтый фон)
                frmDemo.lblInform(intIndex).BackColor = vbYellow
            'Включить контроль "TimeOut" электронной "Кнопки"
                frmDemo.tmrButton(intIndex).Enabled = True
                GoTo ResetController
            End If
            
            'Статус "Relay" - Доступ при "Автоматическом" управлении терминалами
            '  и это Автостоянкa или платноe Предприятиe
        ElseIf intStatusCode = 3 And frmDemo.optAutomatic = True And _
        (intParking(intIndex) = 1 Or intAccess(intIndex) = 1) Then
            'Текущий столбец "Таблицы персон" = 0 (Персона или Терминал)
            gTablePerson.Col = 0
            'Автостоянкa или платноe Предприятиe
            If Trim(gTablePerson.Text) = "DallasParkAcce" Then
            'Если нет свободных мест - игнорировать событие
                If (gParkingPlaceNum <> 0 And gParkFreePlaces = 0) Then _
                GoTo WaitCycle
            'Текущий столбец "Таблицы персон" = 5 (Addr Port Type)
                gTablePerson.Col = 5
            'Если это Временный Клиент Автостоянки или Предприятия
                If Mid(Trim(gTablePerson.Text), 4) <> "CONTR" Then
            'Вызов процедуры-функции формирования ПЕРСОНАЛЬНОГО КОДА
            '  и других данных при АвтоРегистрации Клиентов
            '  (через специальный "Controller" с кнопкой "Dallas")
                    intAutoRegistrCode = _
                    frmDemo.AutoRegDallasButton(gProtocol.strProtocPersonCode, _
                    intIndex, Trim(gTablePerson.Text))
            'Задержать подачу следующей команды на контроллер кнопки "Dallas"
            '  для блокирования двойного ее нажатия или дребезга кнопки
                    frmDemo.tmrRelay.Interval = _
                    frmDemo.tmrButton(intIndex).Interval
                    frmDemo.tmrRelay.Tag = 0
                    frmDemo.tmrRelay.Enabled = True
            'Запрет опроса терминалов
                    gTermContr = 0
            'Цикл ожидания  истечения времени задержки
                    Do While frmDemo.tmrRelay.Tag = 0
            'Обработать возможные события
                        DoEvents
                    Loop
                    frmDemo.tmrRelay.Enabled = False
            'Разрешение опроса терминалов
                    gTermContr = 1
                    GoTo WaitCycle
            'Нажата "Кнопка" терминала типа "CONTR" - обработать событие
                ElseIf Right(Trim(gTablePerson.Text), 5) = "CONTR" Then
            
            ''' ЗАГЛУШКА
            
                    GoTo WaitCycle
                End If
            
            End If
            
            'Доступ через компьютерное управление терминалом, а также
            '  если это Автостоянка, платное Предприятие или
            '  Проходная для Служащих
        ElseIf intStatusCode = 1 And intParking(intIndex) = 0 And intAccess(intIndex) = 0 And _
        intEmploye(intIndex) = 0 Or _
        (intStatusCode = 5 Or intStatusCode = 6 Or intStatusCode = 7) And _
        intParking(intIndex) = 1 Or _
        (intStatusCode = 8 Or intStatusCode = 9 Or intStatusCode = 10) And _
        intAccess(intIndex) = 1 Or _
        (intStatusCode = 0 Or intStatusCode = 1) And _
        intEmploye(intIndex) = 1 Then
            'Анализ дня доступа
            intCalendarCode = CalendarCode()
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            'Анализ времени доступа
            intTimeCode = TimeCode()
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag


 
'##############################################################################
            'Доступ разрешен всегда (в любой день) и Oбщее (Не Индивидуальное) время
            '  доступа разрешеннoе
            If intCalendarCode = 0 And intTimeCode = 0 Then
            'Eсли это не Автостоянка, не платное Предприятие и не Проходная для Служащих
                If intParking(intIndex) = 0 And intAccess(intIndex) = 0 And intEmploye(intIndex) = 0 Then
            'Метка "N_?" - (зеленый фон)
                    frmDemo.lblInform(intIndex).BackColor = vbGreen
            'Очистить приемный буфер порта
                    frmDemo.prtPortC(intIndex).InBufferCount = 0
            'Послать "Controller'y" команду - ОПРОС СОСТОЯНИЯ с собственным адресом
                    frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(224) Or CByte(vntAddr)))
             'Ждать завершения передачи команды ОПРОС СОСТОЯНИЯ
                    Do
                    Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'Ожидание КОДА СОСТОЯНИЯ закрытого терминала от "Controller'a"
                    intScriptCode = ScriptTermClose(intIndex, vntAddr)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            'КОД СОСТОЯНИЯ закрытого терминала поступил
                    If intScriptCode = 0 Then
            'Очистить приемный буфер порта
                        frmDemo.prtPortC(intIndex).InBufferCount = 0
            'Послать "Controller'y" команду - ОТКРЫТЬ ТЕРМИНАЛ с собственным адресом
                        frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(16) Or CByte(vntAddr)))
             'Ждать завершения передачи команды ОТКРЫТЬ ТЕРМИНАЛ
                        Do
                        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'Ожидание КВИТАНЦИИ от "Controller'a"
                        intScriptCode = ScriptOpen(intIndex, vntAddr)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
                    End If
            'Установить признак ЗАНЯТОГО "Controller'a", у которого
            '   ожидается Оранжевый индикатор на считывателе
                    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
                    
            'КОД СОСТОЯНИЯ закрытого терминала и КВИТАНЦИЯ
            '  на команду ОТКРЫТЬ ТЕРМИНАЛ поступили,
                    If intScriptCode = 0 Then
            'Номер считывателя="0" - Разрешен допуск в помещение (на территорию)
                        If vntWork = 0 Then
            'Коррекция примечания из надписи метки "N_?"
                            gProtocol.strProtocReserve = "0/" + Trim(gProtocol.strProtocReserve) + "||" + _
                            CStr(vntReadPortNum) + "Y"
            'Номер считывателя="1" - Разрешен допуск из помещения (с территории)
                        Else
            'Коррекция примечания из надписи метки "N_?"
                            gProtocol.strProtocReserve = "1/" + Trim(gProtocol.strProtocReserve) + "||" + _
                            CStr(vntReadPortNum) + "Y"
                        End If
                    End If
                    
'''            'КОД СОСТОЯНИЯ закрытого терминала или КВИТАНЦИЯ
'''            '  на команду ОТКРЫТЬ ТЕРМИНАЛ не поступили, установлена
'''            '   опция "Шлюз" для терминалов порта
'''                    If intScriptCode <> 0 And gSluice(intIndex) <> 0 Then
'''            'Сбросить все "Controller'ы" порта - СБРОС с групповым адресом
'''                        frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208)))
'''             'Ждать завершения передачи команды СБРОС
'''                        Do
'''                        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
'''                    End If
            'Eсли это Автостоянка,платное Предприятие или Проходная для Служащих
                Else
            'Игнорировать анализ разрешенных дня, терминала и времени доступа - все разрешено
                    intTimeCode = 0
                    intTerminalCode = 0
                    intCalendarCode = 1
                End If
'##############################################################################

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
            'Специальный режим доступа и Oбщее (Не Индивидуальное) время
            '  доступа разрешеннoе
            ElseIf intCalendarCode = 3 And intTimeCode = 0 Then
            'Анализ индивидуального времени доступа
                intTimeCode = IndividualTime(vntAddr, vntReadPortNum)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            'Анализ индивидуального терминала доступа
                intTerminalCode = IndividualTerminal(vntAddr, vntReadPortNum)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            'Анализ индивидуального дня доступа
                intCalendarCode = IndividualCalendar()
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            'Индивидуальные день, время и терминал доступа - (Разрешенные)
            '  и, если это не Автостоянка, не платное Предприятие и не
            '  Проходная для Служащих
                If intCalendarCode = 1 And intTimeCode = 0 _
                And intTerminalCode = 0 And _
                intParking(intIndex) = 0 And intAccess(intIndex) = 0 And _
                intEmploye(intIndex) = 0 Then
            'Метка "N_?" - (зеленый фон)
                    frmDemo.lblInform(intIndex).BackColor = vbGreen
            'Очистить приемный буфер порта
                    frmDemo.prtPortC(intIndex).InBufferCount = 0
            'Послать "Controller'y" команду - ОПРОС СОСТОЯНИЯ с собственным адресом
                    frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(224) Or CByte(vntAddr)))
             'Ждать завершения передачи команды ОПРОС СОСТОЯНИЯ
                    Do
                    Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'Ожидание КОДА СОСТОЯНИЯ закрытого терминала от "Controller'a"
                    intScriptCode = ScriptTermClose(intIndex, vntAddr)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            'КОД СОСТОЯНИЯ закрытого терминала поступил
                    If intScriptCode = 0 Then
            'Очистить приемный буфер порта
                        frmDemo.prtPortC(intIndex).InBufferCount = 0
            'Послать "Controller'y" команду - ОТКРЫТЬ ТЕРМИНАЛ с собственным адресом
                        frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(16) Or CByte(vntAddr)))
             'Ждать завершения передачи команды ОТКРЫТЬ ТЕРМИНАЛ
                        Do
                        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'Ожидание КВИТАНЦИИ от "Controller'a"
                        intScriptCode = ScriptOpen(intIndex, vntAddr)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
                    End If
            'Установить признак ЗАНЯТОГО "Controller'a", у которого
            '   ожидается Оранжевый индикатор на считывателе
                    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
                    
            'КОД СОСТОЯНИЯ закрытого терминала и КВИТАНЦИЯ
            '  на команду ОТКРЫТЬ ТЕРМИНАЛ поступили,
                    If intScriptCode = 0 Then
            'Номер считывателя="0" - Служащий входит в помещение
                        If vntWork = 0 Then
            'Коррекция примечания из надписи метки "N_?"
                            gProtocol.strProtocReserve = "0/" + Trim(gProtocol.strProtocReserve) + "||" + _
                            CStr(vntReadPortNum) + "Y"
            'Номер считывателя="1" - Служащий выходит из помещения
                        Else
            'Коррекция примечания из надписи метки "N_?"
                            gProtocol.strProtocReserve = "1/" + Trim(gProtocol.strProtocReserve) + "||" + _
                            CStr(vntReadPortNum) + "Y"
                        End If
                    End If
                    
'''            'КОД СОСТОЯНИЯ закрытого терминала или КВИТАНЦИЯ
'''            '  на команду ОТКРЫТЬ ТЕРМИНАЛ не поступили, установлена
'''            '   опция "Шлюз" для терминалов порта
'''                    If intScriptCode <> 0 And gSluice(intIndex) <> 0 Then
'''            'Сбросить все "Controller'ы" порта - СБРОС с групповым адресом
'''                        frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208)))
'''             'Ждать завершения передачи команды СБРОС
'''                        Do
'''                        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
'''                    End If
                
                End If
            End If
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX



            'День, время и терминал доступа - (Разрешенные) и
            '  это Автостоянка, платное Предприятие или Проходная для Служащих
            If intCalendarCode = 1 And intTimeCode = 0 And _
            intTerminalCode = 0 And _
            (intParking(intIndex) = 1 Or intAccess(intIndex) = 1 _
            Or intEmploye(intIndex) = 1) Then
            'Метка "N_?" - (зеленый фон)
                frmDemo.lblInform(intIndex).BackColor = vbGreen

            'Очистить приемный буфер порта
                frmDemo.prtPortC(intIndex).InBufferCount = 0
            'Послать "Controller'y" команду - ОПРОС СОСТОЯНИЯ с собственным адресом
                frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(224) Or CByte(vntAddr)))
             'Ждать завершения передачи команды ОПРОС СОСТОЯНИЯ
                Do
                Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'Ожидание КОДА СОСТОЯНИЯ закрытого терминала от "Controller'a"
                intScriptCode = ScriptTermClose(intIndex, vntAddr)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            'КОД СОСТОЯНИЯ закрытого терминала поступил
                If intScriptCode = 0 Then
                
            'Установлена опция - "Автостоянка"
                    If intParking(intIndex) = 1 Then
            'Анализ ячейки "Reserve" в "Таблице персон"
                        intParkingCode = frmTablePerson.AnalysisParking(vntWork)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            'Была некорректная ситуация при анализе ячейки "Reserve"
            '  в "Таблице персон" (просрочена оплата или двойной въезд/выезд)
                        If intParkingCode <> 0 Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
                            frmDemo.BeepSound
            'Номер считывателя="0" - Автомобиль въезжает на Автостоянку
                            If vntWork = 0 Then
            'Метка собщения с предупреждением - на экран
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            'Коррекция примечания из надписи метки "N_?"
                                gProtocol.strProtocReserve = "0/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                CStr(vntReadPortNum) + "N"
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
                                strMessage = "Error Input !!! " + gProtocol.strProtocReserve
            'Номер считывателя="1" - Автомобиль выезжает c Автостоянки
                            Else
            'Метка собщения с предупреждением - на экран
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            'Коррекция примечания из надписи метки "N_?"
                                If intParkingCode = 2 Then
                                    gProtocol.strProtocReserve = "Extra $?"
                                Else
                                    gProtocol.strProtocReserve = "1/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                    CStr(vntReadPortNum) + "N"
                                End If
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
                                strMessage = "Error Output !!! " + gProtocol.strProtocReserve
                            End If
            'Отослать СООБЩЕНИЕ
                            Call frmDemo.SendMessage(strMessage)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            'Имитировать "Плохую" КВИТАНЦИЮ от "Controller'a"
                            intScriptCode = 1
                        Else
            'Очистить приемный буфер порта
                            frmDemo.prtPortC(intIndex).InBufferCount = 0
            'Послать "Controller'y" команду - ОТКРЫТЬ ТЕРМИНАЛ с собственным адресом
                            frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(16) Or CByte(vntAddr)))
             'Ждать завершения передачи команды ОТКРЫТЬ ТЕРМИНАЛ
                            Do
                            Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'Ожидание КВИТАНЦИИ от "Controller'a"
                            intScriptCode = ScriptOpen(intIndex, vntAddr)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
                        End If
            
            'Установлена опция - "Предприятие"
                    ElseIf intAccess(intIndex) = 1 Then
            'Анализ ячейки "Reserve" в "Таблице персон"
                        intAccessCode = frmTablePerson.AnalysisAccess(vntWork)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            'Была некорректная ситуация при анализе ячейки "Reserve"
            '  в "Таблице персон" (просрочена оплата или двойной вход/выход)
                        If intAccessCode <> 0 Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
                            frmDemo.BeepSound
            'Номер считывателя="0" - Посетитель входит на Предприятие
                            If vntWork = 0 Then
            'Метка собщения с предупреждением - на экран
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            'Коррекция примечания из надписи метки "N_?"
                                gProtocol.strProtocReserve = "0/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                CStr(vntReadPortNum) + "N"
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
                                strMessage = "Error Input !!! " + gProtocol.strProtocReserve
            'Номер считывателя="1" - Посетитель выходит с Предприятия
                            Else
            'Метка собщения с предупреждением - на экран
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            'Коррекция примечания из надписи метки "N_?"
                                If intParkingCode = 2 Then
                                    gProtocol.strProtocReserve = "Extra $?"
                                Else
                                    gProtocol.strProtocReserve = "1/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                    CStr(vntReadPortNum) + "N"
                                End If
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
                                strMessage = "Error Output !!! " + gProtocol.strProtocReserve
                            End If
            'Отослать СООБЩЕНИЕ
                            Call frmDemo.SendMessage(strMessage)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            'Имитировать "Плохую" КВИТАНЦИЮ от "Controller'a"
                            intScriptCode = 1
                        Else
            'Очистить приемный буфер порта
                            frmDemo.prtPortC(intIndex).InBufferCount = 0
            'Послать "Controller'y" команду - ОТКРЫТЬ ТЕРМИНАЛ с собственным адресом
                            frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(16) Or CByte(vntAddr)))
             'Ждать завершения передачи команды ОТКРЫТЬ ТЕРМИНАЛ
                            Do
                            Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'Ожидание КВИТАНЦИИ от "Controller'a"
                            intScriptCode = ScriptOpen(intIndex, vntAddr)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
                        End If
            
            'Установлена опция - "Служащий - на проходной"
                    ElseIf intEmploye(intIndex) = 1 Then
            'Анализ ячейки "Name" в "Таблице персон"
                        intEmployeCode = frmTablePerson.AnalysisEmploye(vntWork)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            'Была некорректная ситуация при анализе ячейки "Name"
            '  в "Таблице персон" (двойной вход/выход)
                        If intEmployeCode <> 0 Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
                            frmDemo.BeepSound
            'Номер считывателя="0" - Служащий входит на Предприятие
                            If vntWork = 0 Then
            'Метка собщения с предупреждением - на экран
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            'Коррекция примечания из надписи метки "N_?"
                                gProtocol.strProtocReserve = "0/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                CStr(vntReadPortNum) + "N"
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
                                strMessage = "Error Input !!! " + gProtocol.strProtocReserve
            'Номер считывателя="1" - Служащий выходит с Предприятия
                            Else
            'Метка собщения с предупреждением - на экран
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            'Коррекция примечания из надписи метки "N_?"
                                gProtocol.strProtocReserve = "1/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                CStr(vntReadPortNum) + "N"
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
                                strMessage = "Error Output !!! " + gProtocol.strProtocReserve
                            End If
            'Отослать СООБЩЕНИЕ
                            Call frmDemo.SendMessage(strMessage)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            'Имитировать "Плохую" КВИТАНЦИЮ от "Controller'a"
                            intScriptCode = 1
                        Else
            'Очистить приемный буфер порта
                            frmDemo.prtPortC(intIndex).InBufferCount = 0
            'Послать "Controller'y" команду - ОТКРЫТЬ ТЕРМИНАЛ с собственным адресом
                            frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(16) Or CByte(vntAddr)))
             'Ждать завершения передачи команды ОТКРЫТЬ ТЕРМИНАЛ
                            Do
                            Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'Ожидание КВИТАНЦИИ от "Controller'a"
                            intScriptCode = ScriptOpen(intIndex, vntAddr)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
                        End If
            
            'Опции "Автостоянка", "Посетитель" и Служащий - на проходной" не установлены
                    ElseIf intParking(intIndex) <> 1 And intAccess(intIndex) <> 1 And _
                    intEmploye(intIndex) <> 1 Then
            'Очистить приемный буфер порта
                        frmDemo.prtPortC(intIndex).InBufferCount = 0
            'Послать "Controller'y" команду - ОТКРЫТЬ ТЕРМИНАЛ с собственным адресом
                        frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(16) Or CByte(vntAddr)))
             'Ждать завершения передачи команды ОТКРЫТЬ ТЕРМИНАЛ
                        Do
                        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'Ожидание КВИТАНЦИИ от "Controller'a"
                        intScriptCode = ScriptOpen(intIndex, vntAddr)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
                    End If
            
            'КОД СОСТОЯНИЯ закрытого терминала и КВИТАНЦИЯ
            '  на команду ОТКРЫТЬ ТЕРМИНАЛ поступили, установлена
            '   опция - "Автостоянка"
                    If intScriptCode = 0 And intParking(intIndex) = 1 Then
            
            'Текущий столбец "Таблицы персон" = 0 (Имя)
                        gTablePerson.Col = 0
                        gProtocol.strProtocName = gTablePerson.Text
            'Текущий столбец "Таблицы персон" = 1 (Персональный код)
                        gTablePerson.Col = 1
                        gProtocol.strProtocPersonCode = gTablePerson.Text
            'Текущий столбец "Таблицы персон" = 2 (Статус)
                        gTablePerson.Col = 2
                        gProtocol.strProtocStatus = Trim(gTablePerson.Text)
            'Время
                        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание из надписи метки "N_?"
                        gProtocol.strProtocReserve = frmDemo.lblInform(intIndex).Caption
            
            'Номер считывателя="0" - Автомобиль въехал на Автостоянку
                        If vntWork = 0 Then
            'Коррекция ячейки "Reserve" в "Таблице персон"
                            intParkingCode = frmTablePerson.InputParking(intIndex)
            'Номер считывателя="1" - Автомобиль выехал c Автостоянки
                        Else
            'Коррекция ячейки "Reserve" или исключение строки в "Таблице персон"
                            intParkingCode = frmTablePerson.OutputParking(intIndex, intStatusCode)
                        End If
            'Была некорректная ситуация при коррекции ячейки"Reserve"
            '  в "Таблице персон" (просрочена оплата или двойной въезд/выезд)
                        If intParkingCode <> 0 Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
                            frmDemo.BeepSound
            'Номер считывателя="0" - Автомобиль въезжает на Автостоянку
                            If vntWork = 0 Then
            'Метка собщения с предупреждением - на экран
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            'Коррекция примечания из надписи метки "N_?"
                                gProtocol.strProtocReserve = "0/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                CStr(vntReadPortNum) + "N"
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
                                strMessage = "Error Input !!! " + gProtocol.strProtocReserve
            'Номер считывателя="1" - Автомобиль выезжает c Автостоянки
                            Else
            'Метка собщения с предупреждением - на экран
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            'Коррекция примечания из надписи метки "N_?"
                                gProtocol.strProtocReserve = "1/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                CStr(vntReadPortNum) + "N"
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
                                strMessage = "Error Output !!! " + gProtocol.strProtocReserve
                            End If
            'Отослать СООБЩЕНИЕ
                            Call frmDemo.SendMessage(strMessage)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            'Корректная ситуация
                        Else
            'Номер считывателя="0" - Автомобиль въехал на Автостоянку
                            If vntWork = 0 Then
            'Коррекция поля "Reserve" в "Протоколе"
                                gProtocol.strProtocReserve = "AutoParking || " + "+"
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
                                strMessage = "ParkFreePlaces-1"
            'Номер считывателя="1" - Автомобиль выехал c Автостоянки
                            Else
            'Коррекция поля "Reserve" в "Протоколе"
                                gProtocol.strProtocReserve = "AutoParking || " + "-"
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
                                strMessage = "ParkFreePlaces+1"
                            End If
            'Отослать СООБЩЕНИЕ
                            Call frmDemo.SendMessage(strMessage)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
                       End If
                    
            'КОД СОСТОЯНИЯ закрытого терминала и КВИТАНЦИЯ
            '  на команду ОТКРЫТЬ ТЕРМИНАЛ поступили, установлена
            '   опция - "Посетитель"
                    ElseIf intScriptCode = 0 And intAccess(intIndex) = 1 Then
            
            'Текущий столбец "Таблицы персон" = 0 (Имя)
                        gTablePerson.Col = 0
                        gProtocol.strProtocName = gTablePerson.Text
            'Текущий столбец "Таблицы персон" = 1 (Персональный код)
                        gTablePerson.Col = 1
                        gProtocol.strProtocPersonCode = gTablePerson.Text
            'Текущий столбец "Таблицы персон" = 2 (Статус)
                        gTablePerson.Col = 2
                        gProtocol.strProtocStatus = Trim(gTablePerson.Text)
            'Время
                        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание из надписи метки "N_?"
                        gProtocol.strProtocReserve = frmDemo.lblInform(intIndex).Caption
            
            'Номер считывателя="0" - Посетитель вошел на Предприятие
                        If vntWork = 0 Then
            'Коррекция ячейки "Reserve" в "Таблице персон"
                            intAccessCode = frmTablePerson.InputAccess(intIndex)
            'Номер считывателя="1" - Посетитель вышел с Предприятия
                        Else
            'Коррекция ячейки "Reserve" или исключение строки в "Таблице персон"
                            intAccessCode = frmTablePerson.OutputAccess(intIndex, intStatusCode)
                        End If
            'Была некорректная ситуация при коррекции ячейки"Reserve"
            '  в "Таблице персон" (просрочена оплата или двойной въезд/выезд)
                        If intAccessCode <> 0 Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
                            frmDemo.BeepSound
            'Номер считывателя="0" - Посетитель входит на Предприятие
                            If vntWork = 0 Then
            'Метка сообщения с предупреждением - на экран
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            'Коррекция примечания из надписи метки "N_?"
                                gProtocol.strProtocReserve = "0/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                CStr(vntReadPortNum) + "N"
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
                                strMessage = "Error Input !!! " + gProtocol.strProtocReserve
            'Номер считывателя="1" - Посетитель выходит с Предприятия
                            Else
            'Метка сообщения с предупреждением - на экран
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            'Коррекция примечания из надписи метки "N_?"
                                gProtocol.strProtocReserve = "1/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                CStr(vntReadPortNum) + "N"
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
                                strMessage = "Error Output !!! " + gProtocol.strProtocReserve
                            End If
            'Отослать СООБЩЕНИЕ
                            Call frmDemo.SendMessage(strMessage)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            'Корректная ситуация
                        Else
            'Номер считывателя="0" - Посетитель вошел на Предприятие
                            If vntWork = 0 Then
            'Коррекция поля "Reserve" в "Протоколе"
                                gProtocol.strProtocReserve = "AutoAccess || " + "+"
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
                                strMessage = "AcceFreePlaces-1"
            'Номер считывателя="1" - Посетитель вышел с Предприятия
                            Else
            'Коррекция поля "Reserve" в "Протоколе"
                                gProtocol.strProtocReserve = "AutoAccess || " + "-"
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
                                strMessage = "AcceFreePlaces+1"
                            End If
            'Отослать СООБЩЕНИЕ
                            Call frmDemo.SendMessage(strMessage)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
                        End If
                    
            'КОД СОСТОЯНИЯ закрытого терминала и КВИТАНЦИЯ
            '  на команду ОТКРЫТЬ ТЕРМИНАЛ поступили, установлена
            '   опция - "Служащий - на проходной"
                    ElseIf intScriptCode = 0 And intEmploye(intIndex) = 1 Then
            
            'Текущий столбец "Таблицы персон" = 0 (Имя)
                        gTablePerson.Col = 0
                        gProtocol.strProtocName = gTablePerson.Text
            'Текущий столбец "Таблицы персон" = 1 (Персональный код)
                        gTablePerson.Col = 1
                        gProtocol.strProtocPersonCode = gTablePerson.Text
            'Текущий столбец "Таблицы персон" = 2 (Статус)
                        gTablePerson.Col = 2
                        gProtocol.strProtocStatus = Trim(gTablePerson.Text)
            'Время
                        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание из надписи метки "N_?"
                        gProtocol.strProtocReserve = frmDemo.lblInform(intIndex).Caption
            
            'Номер считывателя="0" - Служащий вошел на Предприятие
                        If vntWork = 0 Then
            'Коррекция ячейки "Name" в "Таблице персон"
                            intEmployeCode = frmTablePerson.InputEmploye(intIndex)
            'Номер считывателя="1" - Служащий вышел с Предприятия
                        Else
            'Коррекция ячейки "Name" в "Таблице персон"
                            intEmployeCode = frmTablePerson.OutputEmploye(intIndex)
                        End If
            'Была некорректная ситуация при коррекции ячейки"Name"
            '  в "Таблице персон" (двойной вход/выход)
                        If intEmployeCode <> 0 Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
                            frmDemo.BeepSound
            'Номер считывателя="0" - Служащий входит на Предприятие
                            If vntWork = 0 Then
            'Метка сообщения с предупреждением - на экран
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            'Коррекция примечания из надписи метки "N_?"
                                gProtocol.strProtocReserve = "0/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                CStr(vntReadPortNum) + "N"
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
                                strMessage = "Error Input !!! " + gProtocol.strProtocReserve
            'Номер считывателя="1" - Служащий выходит с Предприятия
                            Else
            'Метка сообщения с предупреждением - на экран
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            'Коррекция примечания из надписи метки "N_?"
                                gProtocol.strProtocReserve = "1/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                CStr(vntReadPortNum) + "N"
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
                                strMessage = "Error Output !!! " + gProtocol.strProtocReserve
                            End If
            'Отослать СООБЩЕНИЕ
                            Call frmDemo.SendMessage(strMessage)
'Восстановить запомненный ранее номер текущей строки "Таблицы персон"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            'Корректная ситуация
                        Else
            'Номер считывателя="0" - Служащий вошел на Предприятие
                            If vntWork = 0 Then
            'Коррекция поля "Reserve" в "Протоколе"
                                gProtocol.strProtocReserve = "0/" + _
                                CStr(vntAddr) + "/" + _
                                CStr(vntReadPortNum) + "  Input"
            'Номер считывателя="1" - Служащий вышел с Предприятия
                            Else
            'Коррекция поля "Reserve" в "Протоколе"
                                gProtocol.strProtocReserve = "1/" + _
                                CStr(vntAddr) + "/" + _
                                CStr(vntReadPortNum) + "  Output"
                            End If
                        End If
                    
                    End If
                
                End If
            'Установить признак ЗАНЯТОГО "Controller'a", у которого
            '   ожидается Оранжевый индикатор на считывателе
                gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"

'''            'КОД СОСТОЯНИЯ закрытого терминала или КВИТАНЦИЯ
'''            '  на команду ОТКРЫТЬ ТЕРМИНАЛ не поступили, установлена
'''            '   опция "Шлюз" для терминалов порта
'''                If intScriptCode <> 0 And gSluice(intIndex) <> 0 Then
'''            'Сбросить все "Controller'ы" порта - СБРОС с групповым адресом
'''                    frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208)))
'''             'Ждать завершения передачи команды СБРОС
'''                    Do
'''                    Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
'''                End If
            
            'День, время или терминал доступа - (Неразрешенные)
            ElseIf intCalendarCode > 1 Or intTimeCode <> 0 Or _
            intTerminalCode <> 0 Then
            'Метка "N_?" - (белый фон)
                frmDemo.lblInform(intIndex).BackColor = vbWhite
                    
            'Номер считывателя="0"
                If vntWork = 0 Then
            'Коррекция примечания из надписи метки "N_?"
                    gProtocol.strProtocReserve = "0/" + Trim(gProtocol.strProtocReserve) + "||" + _
                    CStr(vntReadPortNum) + "N"
            'Номер считывателя="1"
                Else
            'Коррекция примечания из надписи метки "N_?"
                    gProtocol.strProtocReserve = "1/" + Trim(gProtocol.strProtocReserve) + "||" + _
                    CStr(vntReadPortNum) + "N"
                End If
                    
                GoTo ResetController
            End If
        End If
        
    End If
Protocol:
            'Уменьшить текущее значение атрибута
            '  доступности "Таблицы персон"
    gTablePerson.AccessMinus
            'Записать строку в файл "Таблицы протокола"
    frmDemo.WriteProtocol
            'Выход из процедуры
    Exit Sub
ResetController:
            'Записать строку в файл "Таблицы протокола"
    frmDemo.WriteProtocol
            
'''            'Установлена опция "Шлюз" для терминалов порта
'''    If gSluice(intIndex) <> 0 Then
'''            'Сбросить все "Controller'ы" порта - СБРОС с групповым адресом
'''        frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208)))
'''             'Ждать завершения передачи команды СБРОС
'''        Do
'''        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
'''    End If

            'Цикл ожидания событий
WaitCycle:
            'Уменьшить текущее значение атрибута
            '  доступности "Таблицы персон"
    gTablePerson.AccessMinus

End Sub

            'Анализ статуса доступа
            '   Код возврата: 0 - доступ через локальный терминал;
            '                 1 - доступ через компьютерное управление терминалом;
            '                 2 - "Черный ключ";
            '                 3 - "Relay" или с подтверждением от охраны;
            '                 5 - доступ через компьютерное управление терминалом -
            '                     для Постоянных Клиентов и только для Автостоянки;
            '                 6 - доступ через компьютерное управление терминалом -
            '                     для Временных Клиентов и только для Автостоянки;
            '                 7 - доступ через компьютерное управление терминалом -
            '                     для Бесплатных Клиентов и только для Автостоянки.
            '                 8 - доступ через компьютерное управление терминалом -
            '                     для Постоянных Клиентов и только для Предприятия;
            '                 9 - доступ через компьютерное управление терминалом -
            '                     для Временных Клиентов и только для Предприятия;
            '                10 - доступ через компьютерное управление терминалом -
            '                     для Бесплатных Клиентов и только для Предприятия.
Private Function StatusCode()
            'Статус доступа через локальный терминал
    If Left(gProtocol.strProtocStatus, 2) = "00" Then
        StatusCode = 0
            'Статус доступа через компьютерное управление терминалом
    ElseIf Left(gProtocol.strProtocStatus, 2) = "01" Then
        StatusCode = 1
            'Черный ключ
    ElseIf Left(gProtocol.strProtocStatus, 2) = "02" Then
        StatusCode = 2
            'Статус доступа - "Relay" или с подтверждением от охраны
    ElseIf Left(gProtocol.strProtocStatus, 2) = "03" Then
        StatusCode = 3
            'Статус доступа через компьютерное управление терминалом -
            '  для Постоянных Клиентов и только для Автостоянки
    ElseIf Left(gProtocol.strProtocStatus, 2) = "05" Then
        StatusCode = 5
            'Статус доступа через компьютерное управление терминалом -
            '  для Временных Клиентов и только для Автостоянки
    ElseIf Left(gProtocol.strProtocStatus, 2) = "06" Then
        StatusCode = 6
            'Статус доступа через компьютерное управление терминалом -
            '  для Бесплатных Клиентов и только для Автостоянки
    ElseIf Left(gProtocol.strProtocStatus, 2) = "07" Then
        StatusCode = 7
            'Статус доступа через компьютерное управление терминалом -
            '  для Постоянных Посетителей Предприятия
    ElseIf Left(gProtocol.strProtocStatus, 2) = "08" Then
        StatusCode = 8
            'Статус доступа через компьютерное управление терминалом -
            '  для Временных Посетителей Предприятия
    ElseIf Left(gProtocol.strProtocStatus, 2) = "09" Then
        StatusCode = 9
            'Статус доступа через компьютерное управление терминалом -
            '  для Бесплатных Посетителей Предприятия
    ElseIf Left(gProtocol.strProtocStatus, 2) = "10" Then
        StatusCode = 10
    End If
    
End Function

            'Анализ дня доступа
            '   Код возврата: 0 - доступ в любой день;
            '                 1 - стандартный график доступа - доступный день;
            '                 2 - стандартный график доступа - недоступный день;
            '                 3 - специальный режим доступа.
Private Function CalendarCode()
            'Текущий столбец "Таблицы персон" = 4 (Календарь)
    gTablePerson.Col = 4
            'Доступ в любой день
    If Left(Trim(gTablePerson.Text), 2) = "00" Then
        CalendarCode = 0
            'Стандартный график доступа
    ElseIf Left(Trim(gTablePerson.Text), 2) = "01" Then
            'Доступный день
        If InStr(1, gToday(0), "/") = 0 Then CalendarCode = 1
            'Недоступный день
        If InStr(1, gToday(0), "/*") <> 0 Then CalendarCode = 2
    End If
            'Специальный режим доступа
    If Left(Trim(gTablePerson.Text), 2) = "02" Or _
    InStr(1, gToday(0), "/^") <> 0 Then CalendarCode = 3
                    
End Function

            'Анализ индивидуального дня доступа
            '   Код возврата: 1 - доступный день;
            '                 2 - недоступный день.
Private Function IndividualCalendar()
            'Номер позиции признака "/" в поле "Reservation"
Dim intPosNum As Integer
            'Номер варианта "Таблицы календаря"
Dim intCalendarNum As Integer
            
            'Если возникла ошибка при анализе информации
    On Error GoTo CheckError
            
            'Текущий столбец "Таблицы персон" = 5 (Резервировано)
    gTablePerson.Col = 5
            'Номер позиции первого признака ..."/"
    intPosNum = InStr(1, Trim(gTablePerson.Text), "/")
            'Номер позиции второго признака ..."/"
    intPosNum = InStr(intPosNum + 1, Trim(gTablePerson.Text), "/")
            'Номер варианта "Таблицы календаря" отсутствует - доступный день
    If Len(Trim(gTablePerson.Text)) = intPosNum Then
        IndividualCalendar = 1
    Else
            'Номер варианта "Таблицы календаря"
        intCalendarNum = Mid(Trim(gTablePerson.Text), intPosNum + 1)
            'Если в варианте Даты отсутствует признак "/" - доступный день
        If InStr(1, gToday(intCalendarNum), "/") = 0 Then
            IndividualCalendar = 1
            'Недоступный день
        Else
            IndividualCalendar = 2
        End If
    
    Exit Function
            'Если возникла ошибка при анализе информации
CheckError:
            'День доступа - (Неразрешенный)
        IndividualCalendar = 2
    End If
    
End Function

            'Анализ времени доступа
            '   Код возврата: 0 - разрешенное время доступа;
            '                 1 - неразрешенное время доступа.
Private Function TimeCode()
            'Текущий столбец "Таблицы персон" = 3 (Время)
    gTablePerson.Col = 3
            'Время доступа - (Разрешенное)
    If ((Left(Trim(gTablePerson.Text), 2) < intHour _
    Or Left(Trim(gTablePerson.Text), 2) = intHour _
    And Mid(Trim(gTablePerson.Text), 4, 2) <= intMinute) _
    And (Mid(Trim(gTablePerson.Text), 7, 2) > intHour _
    Or Mid(Trim(gTablePerson.Text), 7, 2) = intHour _
    And Mid(Trim(gTablePerson.Text), 10, 2) >= intMinute)) Or _
((((CInt(Left(Trim(gTablePerson.Text), 2)) * 60 + _
CInt(Mid(Trim(gTablePerson.Text), 4, 2)))) > _
(CInt(Mid(Trim(gTablePerson.Text), 7, 2)) * 60 + _
CInt(Mid(Trim(gTablePerson.Text), 10, 2)))) And _
        (((intHour * 60 + intMinute) >= _
    (CInt(Left(Trim(gTablePerson.Text), 2)) * 60 + _
    CInt(Mid(Trim(gTablePerson.Text), 4, 2))) And _
    (intHour * 60 + intMinute) <= 24 * 60) Or _
        ((intHour * 60 + intMinute) <= _
    (CInt(Mid(Trim(gTablePerson.Text), 7, 2)) * 60 + _
    CInt(Mid(Trim(gTablePerson.Text), 10, 2))) And _
    (intHour * 60 + intMinute) >= 0))) Then

        TimeCode = 0
            
            'Время доступа - (Неразрешенное)
    Else
        TimeCode = 1
    End If
                    
End Function

            'Анализ индивидуального времени доступа
            '   Код возврата: 0 - разрешенное время доступа;
            '                 1 - неразрешенное время доступа.
Private Function IndividualTime(ByVal vntAddr As Variant, ByVal vntReadPortNum _
                                As Variant)
            'Номер позиции признака "/"
Dim intPosNum As Integer
            'Номер варианта "Таблицы времени"
Dim intTimeNum As Integer
            'Номер варианта дополнительной "Таблицы терминалов"
Dim intTerminalNum As Integer
            'Номер варианта дополнительной "Таблицы календаря"
Dim intCalendarNum As Integer
            'Рабочий счетчик
Dim intCount As Integer
            'Рабочий счетчик
Dim intCount_1 As Integer
            
            'Если возникла ошибка при анализе информации
    On Error GoTo CheckError
    
            'Текущий столбец "Таблицы персон" = 5 (Резервировано)
    gTablePerson.Col = 5
            'Номер варианта "Таблицы времени" отсутствует - разрешенное время доступа
    If Left(Trim(gTablePerson.Text), 1) = "/" Then
        IndividualTime = 0
            'Вычислить номер варианта "Таблицы времени"
    Else
            'Номер позиции первого признака ..."/"
        intPosNum = InStr(1, Trim(gTablePerson.Text), "/")
            'Номер варианта "Таблицы времени"
        intTimeNum = Left(Trim(gTablePerson.Text), intPosNum - 1)
            'Время доступа - (Неразрешенное)
        IndividualTime = 1
            'По всем "значащим" столбцам текущей строки массива интервалов
            '  и массива дополнительных (требующих проверки) терминалов
            '  и календарей
        For intCount = 1 To CInt(gInterval(intTimeNum, 0)) - 1 Step 1
            'Время доступа - (Разрешенное)
            If (Left(gInterval(intTimeNum, intCount), 2) < intHour _
            Or Left(gInterval(intTimeNum, intCount), 2) = intHour _
            And Mid(gInterval(intTimeNum, intCount), 4, 2) <= intMinute) _
            And (Mid(gInterval(intTimeNum, intCount), 7, 2) > intHour _
            Or Mid(gInterval(intTimeNum, intCount), 7, 2) = intHour _
            And Mid(gInterval(intTimeNum, intCount), 10, 2) >= intMinute) Then
            'Нет дополнительных (требующих проверки)терминалов или календарей
                If Left(gTerCal(intTimeNum, intCount), 8) = "Interval" Then
                    IndividualTime = 0
                    Exit For
            'Есть дополнительные (требующие проверки)терминалы или календари
                Else
            'Номер позиции признака "/" в массиве дополнительных терминалов
            '  и календарей
                    intPosNum = InStr(1, Trim(gTerCal(intTimeNum, _
                    intCount)), "/")
            'Номер варианта дополнительной "Таблицы терминалов" отсутствует
            '  - разрешенное время доступа
                    If intPosNum = 1 Then
                        IndividualTime = 0
            'Вычислить номер варианта дополнительной "Таблицы терминалов"
                    Else
            'Номер варианта дополнительной "Таблицы терминалов"
                        If intPosNum = 0 Then
                            intTerminalNum = Trim(gTerCal(intTimeNum, _
                            intCount))
                        Else
                            intTerminalNum = Left(Trim(gTerCal(intTimeNum, _
                            intCount)), intPosNum - 1)
                        End If
            'Дополнительный терминал доступа запрещенный
            '  - установка перед проверкой
                        IndividualTime = 1
            'По всем "значащим" столбцам текущей строки массива терминалов
                        For intCount_1 = 1 To _
                        CInt(gAddrPort(intTerminalNum, 0)) - 1 Step 1
            'Дополнительный терминал доступа разрешенный
            '  - установка после проверки
                            If vntAddr = (CByte(Left(gAddrPort(intTerminalNum, _
                            intCount_1), 1) * 16) Or _
                            CByte(Mid(gAddrPort(intTerminalNum, intCount_1), 2, 1))) _
                            And (vntReadPortNum = _
                            CByte(Mid(gAddrPort(intTerminalNum, intCount_1), 3, 1))) Then
                                IndividualTime = 0
                                Exit For
                            End If
                        Next
                    End If
                    
            'В массиве дополнительных терминалов и календарей имеется признак "/"
            '  - необходима проверка дополнительной "Таблицы календаря"
                    If intPosNum <> 0 And IndividualTime = 0 Then
            'Вычислить номер варианта дополнительной "Таблицы календаря"
                        intCalendarNum = Mid(Trim(gTerCal(intTimeNum, intCount)), _
                        intPosNum + 1)
            'Если в варианте Даты имеется признак "/" - недоступный день
                        If InStr(1, gToday(intCalendarNum), "/") <> 0 Then
                            IndividualTime = 1
                        End If
                    End If
            'Все проверки успешные - доступ разрешен
                    If IndividualTime = 0 Then Exit For
                End If
            End If
        Next
    End If
            
    Exit Function
            'Если возникла ошибка при анализе информации
CheckError:
            'Время доступа - (Неразрешенное)
        IndividualTime = 1
                    
End Function

            'Анализ индивидуального терминала доступа
            '   Код возврата: 0 - доступ через данный терминал разрешен;
            '                 1 - доступ через данный терминал запрещен.
Private Function IndividualTerminal(ByVal vntAddr As Variant, ByVal vntReadPortNum _
                                    As Variant)
            'Номер позиции признака "/" в поле "Reservation"
Dim intPosNum As Integer
            'Номер варианта "Таблицы терминалов"
Dim intTerminalNum As Integer
            'Рабочий счетчик
Dim intCount As Integer
            
            'Если возникла ошибка при анализе информации
    On Error GoTo CheckError
            
            'Текущий столбец "Таблицы персон" = 5 (Резервировано)
    gTablePerson.Col = 5
            'Номер позиции первого признака ..."/"
    intPosNum = InStr(1, Trim(gTablePerson.Text), "/")
            'Номер варианта "Таблицы терминалов" отсутствует - доступ разрешен
    If Mid(Trim(gTablePerson.Text), intPosNum + 1, 1) = "/" Then
            'Вычислить номер варианта "Таблицы терминалов"
        IndividualTerminal = 0
    Else
            'Номер варианта "Таблицы терминалов"
        intTerminalNum = Mid(Trim(gTablePerson.Text), intPosNum + 1, _
        InStr(intPosNum + 1, Trim(gTablePerson.Text), "/") - intPosNum - 1)
            'Терминал доступа - (Запрещенный)
        IndividualTerminal = 1
            'По всем "значащим" столбцам текущей строки массива терминалов
        For intCount = 1 To CInt(gAddrPort(intTerminalNum, 0)) - 1 Step 1
            'Терминал доступа - (Разрешенный)
            If vntAddr = (CByte(Left(gAddrPort(intTerminalNum, intCount), 1) * 16) _
            Or CByte(Mid(gAddrPort(intTerminalNum, intCount), 2, 1))) _
            And (vntReadPortNum = CByte(Mid(gAddrPort(intTerminalNum, intCount), 3, 1))) Then
                IndividualTerminal = 0
                Exit For
            End If
        Next
    End If
    
    Exit Function
            'Если возникла ошибка при анализе информации
CheckError:
            'Терминал доступа - (Неразрешенный)
        IndividualTerminal = 1
                    
End Function

            'Ожидание КОДА СОСТОЯНИЯ закрытого терминала от "Controller'a"
            '   в ответ на команду ОПРОС СОСТОЯНИЯ
            '   Код возврата: 0 - поступил КОД СОСТОЯНИЯ закрытого терминала;
            '                 1 - КОД СОСТОЯНИЯ неверный или отсутствует;
            '                 2 - поступил КОД СОСТОЯНИЯ открытого терминала.
Private Function ScriptTermClose(intIndex As Integer, ByVal vntAddr As Variant)
            'Рабочее поле
Dim vntWork As Variant
Dim vntWork1 As Variant
Dim vntWork2 As Variant
            'Номер элемента в массиве "Таблицы терминалов"
Dim intRequest As Integer
            'Запрет опроса терминалов
    gTermContr = 0
            'Вычислить текущий номер порта
    vntWork = frmDemo.prtPortC(intIndex).CommPort
            'Номер текущего элемента
            ' в массиве "Таблицы терминалов",
    intRequest = (vntWork - 2) * 15 + vntAddr
            'Установить признак ЗАНЯТОГО "Controller'a",
            '  от которого ожидается КОД СОСТОЯНИЯ
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "E"
            'Увеличить счетчик "Controller'ов" ЗАНЯТых обслуживанием терминалов
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag + 1
            'Формирование констант для сравнения
    vntWork1 = CByte(16) Or CByte(vntAddr)
    vntWork2 = CByte(32) Or CByte(vntAddr)

            'Включить контроль события TimeOut при ожидании
            '  КОДА СОСТОЯНИЯ от "Controller'a"
    frmDemo.tmrTimeOut(intIndex).Tag = 0
    frmDemo.tmrTimeOut(intIndex).Enabled = True
            'Цикл ожидания КОДа СОСТОЯНИЯ "Controller'a"
    Do While DoEvents()
            'Ждать от "Controller'a" получения КОДА СОСТОЯНИЯ до события TimeOut
        If frmDemo.prtPortC(intIndex).InBufferCount >= 1 Then
            'Полученные данные в приемный буфер для дальнейшей обработки
            vntWork = frmDemo.prtPortC(intIndex).Input
            vntWork = CByte(Asc(vntWork)) And CByte(63)
            'КОД СОСТОЯНИЯ закрытого терминала
            '  от "Controller'a" с соответствующим адресом
            If vntWork = vntWork1 Then
                ScriptTermClose = 0
            'Выход из процедуры
                Exit Do
            'КОД СОСТОЯНИЯ открытого терминала
            '  от "Controller'a" с соответствующим адресом
            ElseIf vntWork = vntWork2 Then
                ScriptTermClose = 2
            'Выход из процедуры
                Exit Do
            'Получены неверные КОД СОСТОЯНИЯ или адрес "Controller'a"
            ElseIf vntWork <> vntWork1 Then
            'Сбросить соответствующий "Controller" - СБРОС с собственным адресом
                frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
             'Ждать завершения передачи команды СБРОС
                Do
                Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'КОД СОСТОЯНИЯ закрытого или открытого
            '  терминала от "Controller'a" не получен
                ScriptTermClose = 1
            'Выход из процедуры
                Exit Do
            End If
            'Произошло событие TimeOut при ожидании КОДА СОСТОЯНИЯ от "Controller'a"
        ElseIf frmDemo.tmrTimeOut(intIndex).Tag <> 0 Then
            'Сбросить соответствующий "Controller" - СБРОС с собственным адресом
            frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
             'Ждать завершения передачи команды СБРОС
            Do
            Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'КОД СОСТОЯНИЯ закрытого или открытого
            '  терминала от "Controller'a" не получен
            ScriptTermClose = 1
            'Выход из процедуры
            Exit Do
        End If
    Loop
            'Выключить контроль события TimeOut
    frmDemo.tmrTimeOut(intIndex).Enabled = False
            'Произошло событие TimeOut или неверный КОД СОСТОЯНИЯ
    If ScriptTermClose = 1 Then
            'Протоколирование события - "TimeOut"
        gProtocol.strProtocName = "Addr=" + Left(gAddrPort(0, intRequest), 2) + _
        " / " + "Port=" + Mid(gAddrPort(0, intRequest), 3, 1)
            'Системный пароль
        gProtocol.strProtocPersonCode = "Command=E/16"
            'Статус
        gProtocol.strProtocStatus = ""
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
        gProtocol.strProtocReserve = "COMMAND TimeOut"
            
            'При установленной опции Логически
            '   Выключить "Controller" из системы
        If intTerminalLogOFF <> 0 Then gAddrPort(0, intRequest) = "00" _
        + Mid(gAddrPort(0, intRequest), 3, 1) + "0"

    End If
            'Сбросить признак ЗАНЯТОГО "Controller'a",
            '  от которого ожидается КОД СОСТОЯНИЯ закрытого терминала
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "0"
            'Уменьшить счетчик "Controller'ов" ЗАНЯТых обслуживанием терминалов
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag - 1
            'Разрешение опроса терминалов
    gTermContr = 1
                    
End Function

            'Ожидание КВИТАНЦИИ от "Controller'a" на команду ОТКРЫТЬ ТЕРМИНАЛ
            '   Код возврата: 0 - КВИТАНЦИЯ поступила;
            '                 1 - КВИТАНЦИЯ отсутствует.
Private Function ScriptOpen(intIndex As Integer, ByVal vntAddr As Variant)
            'Рабочее поле
Dim vntWork As Variant
Dim vntWork1 As Variant
            'Номер элемента в массиве "Таблицы терминалов"
Dim intRequest As Integer
            'Запрет опроса терминалов
    gTermContr = 0
            'Вычислить текущий номер порта
    vntWork = frmDemo.prtPortC(intIndex).CommPort
            'Номер текущего элемента
            ' в массиве "Таблицы терминалов",
    intRequest = (vntWork - 2) * 15 + vntAddr
            'Установить признак ЗАНЯТОГО "Controller'a",
            '  от которого ожидается КВИТАНЦИЯ
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "F"
            'Увеличить счетчик "Controller'ов" ЗАНЯТых обслуживанием терминалов
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag + 1
            'Формирование констант для сравнения
    vntWork1 = CByte(240) Or CByte(vntAddr)

            'Включить контроль события TimeOut при ожидании КВИТАНЦИИ от "Controller'a"
    frmDemo.tmrTimeOut(intIndex).Tag = 0
    frmDemo.tmrTimeOut(intIndex).Enabled = True
            'Цикл опроса "Controller'a"
    Do While DoEvents()
            'Ждать от "Controller'a" получения КВИТАНЦИИ до события TimeOut
        If frmDemo.prtPortC(intIndex).InBufferCount >= 1 Then
            'Полученные данные в приемный буфер для дальнейшей обработки
            vntWork = frmDemo.prtPortC(intIndex).Input
            'КВИТАНЦИЯ от "Controller'a" с соответствующим адресом
            If CByte(Asc(vntWork)) = CByte(vntWork1) Then
                ScriptOpen = 0
            'Bызов процедуры вывода изображения открытого терминала
                PictureTerminalOpen intIndex
            'Выход из процедуры
                Exit Do
            'Получены не КВИТАНЦИЯ или неверный адрес "Controller'a"
            Else
            'Сбросить соответствующий "Controller" - СБРОС с собственным адресом
                frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
             'Ждать завершения передачи команды СБРОС
                Do
                Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'Нет КВИТАНЦИИ
                ScriptOpen = 1
            'Выход из процедуры
                Exit Do
            End If
            'Произошло событие TimeOut при ожидании КВИТАНЦИИ от "Controller'a"
        ElseIf frmDemo.tmrTimeOut(intIndex).Tag <> 0 Then
            'Сбросить соответствующий "Controller" - СБРОС с собственным адресом
            frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
             'Ждать завершения передачи команды СБРОС
            Do
            Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'Произошло событие TimeOut
            ScriptOpen = 1
            'Выход из процедуры
            Exit Do
        End If
    Loop
            'Выключить контроль события TimeOut
    frmDemo.tmrTimeOut(intIndex).Enabled = False
            'Произошло событие TimeOut или нет КВИТАНЦИИ
    If ScriptOpen = 1 Then
            'Вычислить текущий номер порта
        vntWork = frmDemo.prtPortC(intIndex).CommPort
            'Протоколирование события - "TimeOut"
        gProtocol.strProtocName = "Addr=" + Left(gAddrPort(0, intRequest), 2) + _
        " / " + "Port=" + Mid(gAddrPort(0, intRequest), 3, 1)
            'Системный пароль
        gProtocol.strProtocPersonCode = "Command=1/16"
            'Статус
        gProtocol.strProtocStatus = ""
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
        gProtocol.strProtocReserve = "COMMAND TimeOut"
            
            'При установленной опции Логически
            '   Выключить "Controller" из системы
        If intTerminalLogOFF <> 0 Then gAddrPort(0, intRequest) = "00" _
        + Mid(gAddrPort(0, intRequest), 3, 1) + "0"

    End If
            'Сбросить признак ЗАНЯТОГО "Controller'a",
            '  от которого ожидается КВИТАНЦИЯ
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "0"
            'Уменьшить счетчик "Controller'ов" ЗАНЯТых обслуживанием терминалов
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag - 1
            'Разрешение опроса терминалов
    gTermContr = 1
                    
End Function
            
            'Начальная последовательность закрытия терминала
Private Sub InitialCloseTerminal(intIndex As Integer, intRequest As Integer)
            
            'Код возврата при ожидании КОДА СОСТОЯНИЯ или КВИТАНЦИИ
Dim intScriptCode As Integer
            'Адрес контроллера
Dim vntAddr As Variant
    
            ' "Controller" ЛОГИЧЕСКИ выключен из системы
    If Mid(gAddrPort(0, intRequest), 1, 2) = "00" Then GoTo WaitCycle

        'Сформировать адрес "Controller'a", ЗАНЯТого обслуживанием терминала
    vntAddr = CByte(Left(gAddrPort(0, intRequest), 1) * 16) _
    Or CByte(Mid(gAddrPort(0, intRequest), 2, 1))
            'Очистить приемный буфер порта
    frmDemo.prtPortC(intIndex).InBufferCount = 0
            'Послать "Controller'y" команду - ОПРОС СОСТОЯНИЯ с собственным адресом
    frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(224) Or CByte(vntAddr)))
             'Ждать завершения передачи команды ОПРОС СОСТОЯНИЯ
    Do
    Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'Ожидание КОДА СОСТОЯНИЯ открытого терминала от "Controller'a"
    intScriptCode = ScriptTermOpen(intIndex, vntAddr)
            'Цикл ожидания событий
WaitCycle:

End Sub

            'Ожидание КОДА СОСТОЯНИЯ открытого терминала от "Controller'a"
            '   в ответ на команду ОПРОС СОСТОЯНИЯ
            '   Код возврата: 0 - поступил КОД СОСТОЯНИЯ открытого терминала;
            '                 1 - КОД СОСТОЯНИЯ неверный или отсутствует;
            '                 2 - поступил КОД СОСТОЯНИЯ закрытого терминала.
Private Function ScriptTermOpen(intIndex As Integer, ByVal vntAddr As Variant)
            'Рабочее поле
Dim vntWork As Variant
Dim vntWork1 As Variant
Dim vntWork2 As Variant
            'Номер элемента в массиве "Таблицы терминалов"
Dim intRequest As Integer
            'Запрет опроса терминалов
    gTermContr = 0
            'Вычислить текущий номер порта
    vntWork = frmDemo.prtPortC(intIndex).CommPort
            'Номер текущего элемента
            ' в массиве "Таблицы терминалов",
    intRequest = (vntWork - 2) * 15 + vntAddr
            'Установить признак ЗАНЯТОГО "Controller'a",
            '  от которого ожидается КОД СОСТОЯНИЯ
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "E"
            'Увеличить счетчик "Controller'ов" ЗАНЯТых обслуживанием терминалов
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag + 1
            'Формирование констант для сравнения
    vntWork1 = CByte(32) Or CByte(vntAddr)
    vntWork2 = CByte(16) Or CByte(vntAddr)

            'Включить контроль события TimeOut при ожидании
            '  КОДА СОСТОЯНИЯ от "Controller'a"
    frmDemo.tmrTimeOut(intIndex).Tag = 0
    frmDemo.tmrTimeOut(intIndex).Enabled = True
            'Цикл опроса "Controller'a"
    Do While DoEvents()
            'Ждать от "Controller'a" получения КОДА СОСТОЯНИЯ до события TimeOut
        If frmDemo.prtPortC(intIndex).InBufferCount >= 1 Then
            'Полученные данные в приемный буфер для дальнейшей обработки
            vntWork = frmDemo.prtPortC(intIndex).Input
            vntWork = CByte(Asc(vntWork)) And CByte(63)
            'КОД СОСТОЯНИЯ открытого терминала
            '  от "Controller'a" с соответствующим адресом
            If vntWork = vntWork1 Then
                ScriptTermOpen = 0
            'Выход из процедуры
                Exit Do
            'КОД СОСТОЯНИЯ закрытого терминала
            '  от "Controller'a" с соответствующим адресом
            ElseIf vntWork = vntWork2 Then
                ScriptTermOpen = 2
            'Выход из процедуры
                Exit Do
            'Получены неверные КОД СОСТОЯНИЯ или адрес "Controller'a"
            ElseIf vntWork <> vntWork1 Then
            'Сбросить соответствующий "Controller" - СБРОС с собственным адресом
                frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
             'Ждать завершения передачи команды СБРОС
                Do
                Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'КОД СОСТОЯНИЯ закрытого или открытого
            '  терминала от "Controller'a" не получен
                ScriptTermOpen = 1
                Exit Do
            End If
            'Произошло событие TimeOut при ожидании КОДА СОСТОЯНИЯ от "Controller'a"
        ElseIf frmDemo.tmrTimeOut(intIndex).Tag <> 0 Then
            'Сбросить соответствующий "Controller" - СБРОС с собственным адресом
            frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
             'Ждать завершения передачи команды СБРОС
            Do
            Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'Произошло событие TimeOut
            ScriptTermOpen = 1
            Exit Do
        End If
    Loop
            'Выключить контроль события TimeOut
    frmDemo.tmrTimeOut(intIndex).Enabled = False
            'Произошло событие TimeOut или неверный КОД СОСТОЯНИЯ
    If ScriptTermOpen = 1 Then
            'Протоколирование события - "TimeOut"
        gProtocol.strProtocName = "Addr=" + Left(gAddrPort(0, intRequest), 2) + _
        " / " + "Port=" + Mid(gAddrPort(0, intRequest), 3, 1)
            'Системный пароль
        gProtocol.strProtocPersonCode = "Command=E/16"
            'Статус
        gProtocol.strProtocStatus = ""
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
        gProtocol.strProtocReserve = "COMMAND TimeOut"
''            'Записать строку в файл "Таблицы протокола"
''        frmDemo.WriteProtocol
            
            'При установленной опции Логически
            '   Выключить "Controller" из системы
        If intTerminalLogOFF <> 0 Then gAddrPort(0, intRequest) = "00" _
        + Mid(gAddrPort(0, intRequest), 3, 1) + "0"

    End If
            'Сбросить признак ЗАНЯТОГО "Controller'a",
            '  от которого ожидается КОД СОСТОЯНИЯ открытого терминала
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "0"
            'Уменьшить счетчик "Controller'ов" ЗАНЯТых обслуживанием терминалов
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag - 1
            'Разрешение опроса терминалов
    gTermContr = 1
                    
End Function

            'Ожидание КВИТАНЦИИ от "Controller'a" на команду ЗАКРЫТЬ ТЕРМИНАЛ
            '   Код возврата: 0 - КВИТАНЦИЯ поступила;
            '                 1 - КВИТАНЦИЯ отсутствует.
Private Function ScriptClose(intIndex As Integer, ByVal vntAddr As Variant)
            'Рабочее поле
Dim vntWork As Variant
Dim vntWork1 As Variant
            'Номер элемента в массиве "Таблицы терминалов"
Dim intRequest As Integer
            'Запрет опроса терминалов
    gTermContr = 0
            'Вычислить текущий номер порта
    vntWork = frmDemo.prtPortC(intIndex).CommPort
            'Номер текущего элемента
            ' в массиве "Таблицы терминалов",
    intRequest = (vntWork - 2) * 15 + vntAddr
            'Установить признак ЗАНЯТОГО "Controller'a",
            '  от которого ожидается КВИТАНЦИЯ
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "F"
            'Увеличить счетчик "Controller'ов" ЗАНЯТых обслуживанием терминалов
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag + 1
            'Формирование константы для сравнения
    vntWork1 = CByte(240) Or CByte(vntAddr)

            'Включить контроль события TimeOut при ожидании КВИТАНЦИИ от "Controller'a"
    frmDemo.tmrTimeOut(intIndex).Tag = 0
    frmDemo.tmrTimeOut(intIndex).Enabled = True
            'Цикл опроса "Controller'a"
    Do While DoEvents()
            'Ждать от "Controller'a" получения КВИТАНЦИИ до события TimeOut
        If frmDemo.prtPortC(intIndex).InBufferCount >= 1 Then
            'Полученные данные в приемный буфер для дальнейшей обработки
            vntWork = frmDemo.prtPortC(intIndex).Input
            'КВИТАНЦИЯ от "Controller'a" с соответствующим адресом
            If CByte(Asc(vntWork)) = CByte(vntWork1) Then
                ScriptClose = 0
            'Выход из процедуры
                Exit Do
            'Получены не КВИТАНЦИЯ или неверный адрес "Controller'a"
            Else
            'Сбросить соответствующий "Controller" - СБРОС с собственным адресом
                frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
             'Ждать завершения передачи команды СБРОС
                Do
                Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'КВИТАНЦИЯ от "Controller'a" не получена
                ScriptClose = 1
            'Выход из процедуры
                Exit Do
            End If
            'Произошло событие TimeOut при ожидании КВИТАНЦИИ от "Controller'a"
        ElseIf frmDemo.tmrTimeOut(intIndex).Tag <> 0 Then
            'Сбросить соответствующий "Controller" - СБРОС с собственным адресом
            frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
             'Ждать завершения передачи команды СБРОС
            Do
            Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'КВИТАНЦИЯ от "Controller'a" не получена
            ScriptClose = 1
            'Выход из процедуры
            Exit Do
        End If
    Loop
            'Выключить контроль события TimeOut
    frmDemo.tmrTimeOut(intIndex).Enabled = False
            'Произошло событие TimeOut или нет КВИТАНЦИИ
    If ScriptClose = 1 Then
            'Вычислить текущий номер порта
        vntWork = frmDemo.prtPortC(intIndex).CommPort
            'Протоколирование события - "TimeOut"
        gProtocol.strProtocName = "Addr=" + Left(gAddrPort(0, intRequest), 2) + _
        " / " + "Port=" + Mid(gAddrPort(0, intRequest), 3, 1)
            'Системный пароль
        gProtocol.strProtocPersonCode = "Command=A/16"
            'Статус
        gProtocol.strProtocStatus = ""
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
        gProtocol.strProtocReserve = "COMMAND TimeOut"
''            'Записать строку в файл "Таблицы протокола"
''        frmDemo.WriteProtocol
            
            'При установленной опции Логически
            '   Выключить "Controller" из системы
        If intTerminalLogOFF <> 0 Then gAddrPort(0, intRequest) = "00" _
        + Mid(gAddrPort(0, intRequest), 3, 1) + "0"

    End If
            'Восстановить признак ЗАНЯТОГО "Controller'a",
            '  от которого ожидается КВИТАНЦИЯ
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "A"
            'Уменьшить счетчик "Controller'ов" ЗАНЯТых обслуживанием терминалов
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag - 1
            'Разрешение опроса терминалов
    gTermContr = 1
                    
End Function
            
            'Последовательность ожидания открытия терминала
Private Sub WaitOpenTerminal(intIndex As Integer, intRequest As Integer)
            'Код возврата при ожидании КОДА СОСТОЯНИЯ
Dim intScriptCode As Integer
            'Адрес контроллера
Dim vntAddr As Variant
    
            'Запрет опроса терминалов
    gTermContr = 0
            
            'Установить признак ЗАНЯТОГО "Controller'a",
            '  от которого ожидается КОД СОСТОЯНИЯ
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "E"
            'Увеличить счетчик "Controller'ов" ЗАНЯТых обслуживанием терминалов
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag + 1
        
        'Сформировать адрес "Controller'a", ЗАНЯТого обслуживанием терминала
    vntAddr = CByte(Left(gAddrPort(0, intRequest), 1) * 16) _
    Or CByte(Mid(gAddrPort(0, intRequest), 2, 1))
            
            'Включить контроль события TimeOut при ожидании
            '  КОДА СОСТОЯНИЯ от "Controller'a"
    frmDemo.tmrButton(intIndex).Tag = 0
    frmDemo.tmrButton(intIndex).Enabled = True
            'Цикл ожидания кода возврата (закрытого терминала)
    Do While DoEvents()
            
            'Очистить приемный буфер порта
        frmDemo.prtPortC(intIndex).InBufferCount = 0
            'Послать "Controller'y" команду - ОПРОС СОСТОЯНИЯ с собственным адресом
        frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(224) Or CByte(vntAddr)))
             'Ждать завершения передачи команды ОПРОС СОСТОЯНИЯ
        Do
        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'Ожидание кода возврата (открытого терминала)
        intScriptCode = TerminalOpen(intIndex, vntAddr)
            
            'Произошло событие TimeOut при ожидании КОДА СОСТОЯНИЯ от "Controller'a"
        If frmDemo.tmrButton(intIndex).Tag <> 0 Then
            'Выход из процедуры
            Exit Do
            'Kод возврата открытого терминала поступил
        ElseIf intScriptCode = 0 Then
            'Выход из процедуры
            Exit Do
            'Kод возврата открытого терминала еще не поступил
            '  - продолжать ждать кода возврата открытого терминала
        ElseIf intScriptCode <> 0 Then
            
        End If
    Loop
            'Выключить контроль события TimeOut
    frmDemo.tmrButton(intIndex).Enabled = False
            
            'Сбросить признак ЗАНЯТОГО "Controller'a",
            '  от которого ожидается КОД СОСТОЯНИЯ открытого терминала
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "0"
            'Уменьшить счетчик "Controller'ов" ЗАНЯТых обслуживанием терминалов
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag - 1
            
            'Вызов процедуры вывода изображения открытого терминала
    PictureTerminalOpen intIndex
            
            'Разрешение опроса терминалов
    gTermContr = 1
    
End Sub
            
            'Последовательность ожидания закрытия терминала
Private Sub WaitCloseTerminal(intIndex As Integer, intRequest As Integer)
            'Код возврата при ожидании КОДА СОСТОЯНИЯ
Dim intScriptCode As Integer
            'Адрес контроллера
Dim vntAddr As Variant
    
            'Запрет опроса терминалов
    gTermContr = 0
            
            'Установить признак ЗАНЯТОГО "Controller'a",
            '  от которого ожидается КОД СОСТОЯНИЯ
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "E"
            'Увеличить счетчик "Controller'ов" ЗАНЯТых обслуживанием терминалов
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag + 1
        
        'Сформировать адрес "Controller'a", ЗАНЯТого обслуживанием терминала
    vntAddr = CByte(Left(gAddrPort(0, intRequest), 1) * 16) _
    Or CByte(Mid(gAddrPort(0, intRequest), 2, 1))
            
            'Включить контроль события TimeOut при ожидании
            '  КОДА СОСТОЯНИЯ от "Controller'a"
    frmDemo.tmrButton(intIndex).Tag = 0
    frmDemo.tmrButton(intIndex).Enabled = True
            'Цикл ожидания кода возврата (закрытого терминала)
    Do While DoEvents()
            
            'Очистить приемный буфер порта
        frmDemo.prtPortC(intIndex).InBufferCount = 0
            'Послать "Controller'y" команду - ОПРОС СОСТОЯНИЯ с собственным адресом
        frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(224) Or CByte(vntAddr)))
             'Ждать завершения передачи команды ОПРОС СОСТОЯНИЯ
        Do
        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'Ожидание кода возврата (закрытого терминала)
        intScriptCode = TerminalClose(intIndex, vntAddr)
            
            'Произошло событие TimeOut при ожидании КОДА СОСТОЯНИЯ от "Controller'a"
        If frmDemo.tmrButton(intIndex).Tag <> 0 Then
            'Выход из процедуры
            Exit Do
            'Kод возврата закрытого терминала поступил
        ElseIf intScriptCode = 0 Then
            'Выход из процедуры
            Exit Do
            'Kод возврата закрытого терминала еще не поступил
            '  - продолжать ждать кода возврата закрытого терминала
        ElseIf intScriptCode <> 0 Then
                
        End If
    Loop
            'Выключить контроль события TimeOut
    frmDemo.tmrButton(intIndex).Enabled = False
            
            'Сбросить признак ЗАНЯТОГО "Controller'a",
            '  от которого ожидается КОД СОСТОЯНИЯ закрытого терминала
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "0"
            'Уменьшить счетчик "Controller'ов" ЗАНЯТых обслуживанием терминалов
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag - 1
            
            'Вызов процедуры вывода изображения закрытого терминала
    PictureTerminalClose intIndex
            
            'Разрешение опроса терминалов
    gTermContr = 1
    
End Sub
            
            ' "Последовательность открытия терминала"
            '    от электронной "Кнопки"
Private Sub ButtonOpenTerminal(intIndex As Integer, intRequest As Integer)
            'Адрес "Controller'a"
Dim vntAddr As Variant
            'Код возврата при ожидании КВИТАНЦИИ или КОДА СОСТОЯНИЯ
Dim intScriptCode As Integer

            'Определить адрес "Controller'a"
    vntAddr = CByte(Left(gAddrPort(0, intRequest), 1) * 16) _
    Or CByte(Mid(gAddrPort(0, intRequest), 2, 1))
            
            'Очистить приемный буфер порта
    frmDemo.prtPortC(intIndex).InBufferCount = 0
            'Послать "Controller'y" команду - ОПРОС СОСТОЯНИЯ с собственным адресом
    frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(224) Or CByte(vntAddr)))
             'Ждать завершения передачи команды ОПРОС СОСТОЯНИЯ
    Do
    Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'Ожидание КОДА СОСТОЯНИЯ закрытого терминала от "Controller'a"
    intScriptCode = ScriptTermClose(intIndex, vntAddr)
            'КОД СОСТОЯНИЯ закрытого терминала поступил
    If intScriptCode = 0 Then
            'Очистить приемный буфер порта
        frmDemo.prtPortC(intIndex).InBufferCount = 0
            'Послать "Controller'y" команду - ОТКРЫТЬ ТЕРМИНАЛ с собственным адресом
        frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(16) Or CByte(vntAddr)))
             'Ждать завершения передачи команды ОТКРЫТЬ ТЕРМИНАЛ
        Do
        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'Ожидание КВИТАНЦИИ от "Controller'a"
        intScriptCode = ScriptOpen(intIndex, vntAddr)
    End If
            'Установить признак ЗАНЯТОГО "Controller'a", у которого
            '   ожидается Оранжевый индикатор на считывателе
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            'Уменьшить счетчик "Controller'ов", ЗАНЯТых обслуживанием терминалов
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag - 1
            
End Sub
            
            'Процедура вывода изображения открытого терминала
Private Sub PictureTerminalOpen(intIndex As Integer)
        frmDemo.imgViewClose(intIndex).Visible = False
        frmDemo.imgViewOpen(intIndex).Visible = True
            'Включение таймера электронной "Кнопки" (используется
            '  для вывода изображения закрытоко терминала)
        frmDemo.tmrButton(intIndex).Tag = 0
        frmDemo.tmrButton(intIndex).Enabled = True

End Sub
            
            'Процедура вывода изображения закрытого терминала
Private Sub PictureTerminalClose(intIndex As Integer)
        frmDemo.imgViewOpen(intIndex).Visible = False
        frmDemo.imgViewClose(intIndex).Visible = True

End Sub

            'Ожидание КОДА СОСТОЯНИЯ открытого терминала от "Controller'a"
            '   в ответ на команду ОПРОС СОСТОЯНИЯ
            '   Код возврата: 0 - поступил КОД СОСТОЯНИЯ открытого терминала;
            '                 1 - КОД СОСТОЯНИЯ неверный или отсутствует.
Private Function TerminalOpen(intIndex As Integer, ByVal vntAddr As Variant)
            'Рабочее поле
Dim vntWork As Variant
Dim vntWork1 As Variant
Dim vntWork2 As Variant
            
            'Формирование констант для сравнения
    vntWork1 = CByte(16) Or CByte(vntAddr)
    vntWork2 = CByte(0) Or CByte(vntAddr)

            'Включить контроль события TimeOut при ожидании
            '  КОДА СОСТОЯНИЯ от "Controller'a"
    frmDemo.tmrRelay.Interval = frmDemo.tmrTimeOut(intIndex).Interval
    frmDemo.tmrRelay.Tag = 0
    frmDemo.tmrRelay.Enabled = True
            'Цикл ожидания КОДа СОСТОЯНИЯ "Controller'a"
    Do While DoEvents()
            'Ждать от "Controller'a" получения КОДА СОСТОЯНИЯ до события TimeOut
        If frmDemo.prtPortC(intIndex).InBufferCount >= 1 Then
            'Полученные данные в приемный буфер для дальнейшей обработки
            vntWork = frmDemo.prtPortC(intIndex).Input
            vntWork = CByte(Asc(vntWork)) And CByte(31)
            'КОД СОСТОЯНИЯ открытого терминала
            '  от "Controller'a" с соответствующим адресом
            If vntWork = vntWork2 Then
                TerminalOpen = 0
            'Выход из процедуры
                Exit Do
            'КОД СОСТОЯНИЯ закрытого терминала от "Controller'a" с соответствующим
            '  адресом - продолжать ждать КОДа СОСТОЯНИЯ открытого терминала
            ElseIf vntWork = vntWork1 Then
            
                TerminalOpen = 2 'Нигде далее не используется
                
            'Получены неверные КОД СОСТОЯНИЯ или адрес "Controller'a"
            ElseIf vntWork <> vntWork1 Then
            'Сбросить соответствующий "Controller" - СБРОС с собственным адресом
                frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
             'Ждать завершения передачи команды СБРОС
                Do
                Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'КОД СОСТОЯНИЯ закрытого или открытого
            '  терминала от "Controller'a" не получен
                TerminalOpen = 1
            'Выход из процедуры
                Exit Do
            End If
            'Произошло событие TimeOut при ожидании КОДА СОСТОЯНИЯ от "Controller'a"
        ElseIf frmDemo.tmrRelay.Tag <> 0 Then
            'КОД СОСТОЯНИЯ открытого терминала от "Controller'a" не получен
            TerminalOpen = 1
            'Выход из процедуры
            Exit Do
        End If
    Loop
            'Выключить контроль события TimeOut
    frmDemo.tmrRelay.Enabled = False
                    
End Function

            'Ожидание КОДА СОСТОЯНИЯ закрытого терминала от "Controller'a"
            '   в ответ на команду ОПРОС СОСТОЯНИЯ
            '   Код возврата: 0 - поступил КОД СОСТОЯНИЯ закрытого терминала;
            '                 1 - КОД СОСТОЯНИЯ неверный или отсутствует.
Private Function TerminalClose(intIndex As Integer, ByVal vntAddr As Variant)
            'Рабочее поле
Dim vntWork As Variant
Dim vntWork1 As Variant
Dim vntWork2 As Variant
            
            'Формирование констант для сравнения
    vntWork1 = CByte(16) Or CByte(vntAddr)
    vntWork2 = CByte(0) Or CByte(vntAddr)

            'Включить контроль события TimeOut при ожидании
            '  КОДА СОСТОЯНИЯ от "Controller'a"
    frmDemo.tmrRelay.Interval = frmDemo.tmrTimeOut(intIndex).Interval
    frmDemo.tmrRelay.Tag = 0
    frmDemo.tmrRelay.Enabled = True
            'Цикл ожидания КОДа СОСТОЯНИЯ "Controller'a"
    Do While DoEvents()
            'Ждать от "Controller'a" получения КОДА СОСТОЯНИЯ до события TimeOut
        If frmDemo.prtPortC(intIndex).InBufferCount >= 1 Then
            'Полученные данные в приемный буфер для дальнейшей обработки
            vntWork = frmDemo.prtPortC(intIndex).Input
            vntWork = CByte(Asc(vntWork)) And CByte(31)
            'КОД СОСТОЯНИЯ закрытого терминала
            '  от "Controller'a" с соответствующим адресом
            If vntWork = vntWork1 Then
                TerminalClose = 0
            'Выход из процедуры
                Exit Do
            'КОД СОСТОЯНИЯ открытого терминала от "Controller'a" с соответствующим
            '  адресом - продолжать ждать КОДа СОСТОЯНИЯ закрытого терминала
            ElseIf vntWork = vntWork2 Then
            
                TerminalClose = 2 'Нигде далее не используется
                
            'Получены неверные КОД СОСТОЯНИЯ или адрес "Controller'a"
            ElseIf vntWork <> vntWork1 Then
            'Сбросить соответствующий "Controller" - СБРОС с собственным адресом
                frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
             'Ждать завершения передачи команды СБРОС
                Do
                Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            'КОД СОСТОЯНИЯ закрытого или открытого
            '  терминала от "Controller'a" не получен
                TerminalClose = 1
            'Выход из процедуры
                Exit Do
            End If
            'Произошло событие TimeOut при ожидании КОДА СОСТОЯНИЯ от "Controller'a"
        ElseIf frmDemo.tmrRelay.Tag <> 0 Then
            'КОД СОСТОЯНИЯ закрытого терминала от "Controller'a" не получен
            TerminalClose = 1
            'Выход из процедуры
            Exit Do
        End If
    Loop
            'Выключить контроль события TimeOut
    frmDemo.tmrRelay.Enabled = False
            
End Function

            'Копирование "Таблицы персон" из "Host Computer'a"
Private Sub TablePersonCopy()
            'Текущий номер строки таблицы "TableSystem"
            '   в "Host Computer'e"
Dim intRowNum As Integer
            'Полное имя копируемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Полное имя файла-копии (с указанием "пути" к нему)
Dim strCopyFileName As String
            'Полное имя папки-файла "Host Computer'a" (с указанием "пути" к ней)
Dim strPathFolderName As String
            'Объект "FileSystemObject" - "Файловая Система"
Dim FSO As Variant
            
            'Текущий столбец "Системной таблицы" = 0 (Имя)
    frmTableSystem.grdTableSystem.Col = 0
            'Цикл по всем нефиксированным строкам "Системной таблицы"
    For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            'Текущая строка "Системной таблицы"
        frmTableSystem.grdTableSystem.Row = intRowNum
            'Анализ признака копирования "Таблицы персон" из "Host Computer'a"
        If Trim(frmTableSystem.grdTableSystem.Text) = _
        "CopyTablePerson" Then
            'Текущий столбец "Системной таблицы" = 1
            frmTableSystem.grdTableSystem.Col = 1
            'Требуется копирование "Таблицы персон" из "Host Computer'a"
            If Mid(Trim(frmTableSystem.grdTableSystem.Text), 2, 2) = ":\" Then
            'Полное имя папки-файла "Host Computer'a" (с указанием "пути" к ней)
                strPathFolderName = Trim(frmTableSystem.grdTableSystem.Text)
                Exit For
            'Не требуется копирование "Таблицы персон" из "Host Computer'a"
            Else
                Exit Sub
            End If
        End If
    Next
            'Отсутствует строка "Системной таблицы", хранящая признак
            '  необходимости копирования "Таблицы персон" из "Host Computer'a"
    If intRowNum = frmTableSystem.grdTableSystem.Rows Then Exit Sub
            
            'Создать объект "FSO" - "Файловая система"
    Set FSO = CreateObject("Scripting.FileSystemObject")
            'Определить действительный "путь" к каталогу выполняемой программы
    strCopyFileName = App.Path
    If Right(strCopyFileName, 1) <> "\" Then
            'Полное имя папки для файла-копии (с указанием "пути" к ней)
        strCopyFileName = strCopyFileName + "\"
    End If
    
            'Проверка существования папки-файла "Host Computer'a"
    On Error GoTo UnExist
            'Папка-файл имеется - продолжить
    If (FSO.FolderExists(strPathFolderName)) Then
            'Полное имя копируемого файла "Host Computer'a"
            '  (с указанием "пути" к нему)
        strPathFileName = strPathFolderName + "\" + "TablePerson.dat"

        If (FSO.FileExists(strPathFileName)) Then
            'Файл имеется - копирование Архива в "Host Computer"
            FSO.CopyFile strPathFileName, strCopyFileName
        End If
    End If

UnExist:
    On Error GoTo 0

End Sub


