Attribute VB_Name = "Config"
'Все переменные необходимо объявлять заранее
Option Explicit

Global Const isDebug = true

Global Const gdUpLevel = 200 * 0.0981    'Предел давления для заправки

Global Const gdRashAkkEnd = 65    'Нижний расход по которому отсекается поток от аккумуляторов

'Определение констант
'Global Const ggMinPress = 197    ' Минимальное давление в аккумуляторах

Global Const conHwndTopmost = -1
Global Const conSwpNoActivate = &H10
Global Const conSwpShowWindow = &H40

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long


'Объявление пользовательского класса
'Данные датчика + описание
Public Type Sensor
    Data            As Integer
    Note            As String
End Type

'Для расхода газа
Public gdИР1        As Double
Public gdИР2        As Double


'Для заправки
Public gbOnlyAkk    As Boolean    'Флаг заправки только от аккумуляторов
'Только когда не работает ДВС
Public giStage      As Integer    'Этапы заправки
Public giStage1     As Integer    'ПодЭтапы предпускового этапа
Public giStage2     As Integer    'ПодЭтапы этапа Заправки
Public gbFrmShow    As Boolean
Public gbCmdStart   As Boolean    'Вид кнопки Пуск\Заправка
Public gbAkkum      As Boolean
Public giTrigger    As Integer    ' Для формы Запроса
Public gsMsg        As String
Public Car          As Integer
Public gbЗаправка   As Boolean    'Показывает когда нужно в главном цикле считать расход
'газа на заправку одного автомобиля
Public gdРасход1    As Double    'Расход на заправку автомобиля
Public giMainРасход As Integer    'Флаг для подсчета расхода всего газа:
' 1 - добавляем
' 0 - ничего
'-1 - отнимаем

Public gdK          As Double    'Поправочный коэффициент

' Переменные для усреднения
Public glAver       As Long    ' размер цикла дла усреднения
Public glCounter    As Long
Public sum(31)      As Double


Public gbRunDVS     As Boolean





Public FileHandle   As Integer    'Идентификатор файла с описаниями
Public MotorCount   As Long    'Счетчик моторесурсов
Public GMC          As Long    'глобальный счетчик





Public gDateRec     As Date    'Дата последней записи
Public giCountZ     As Integer    'Счетчик заправок
Public giRealCountZ As Integer    'Реальный счетчик заправок

Public gbDontStat   As Boolean    ' флаг заправки (работы)


'Аварийные ситуации
Public gbStopAGNKS  As Boolean    ' флаг Останова АГНКС
Public gbFireDVS    As Boolean    'пожар в отсеке ДВС
Public gbFireTech   As Boolean    'пожар в тех. отсеке


'Болванки
Public gdTime       As Double    'Время заправки


Public giDVS        As Integer
Public gdPlot       As Double