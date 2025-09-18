Imports System.IO

Module Module1
    Private WorkPath As String = ""
    Private SubFolders As Boolean = True
    Private WhiteDays() As Integer = New Integer() {}
    Private MinAge As Integer = 1
    Private FileMask As String = ""
    Private ViewOnly As Boolean = False
    Private ReadOnly Days() As String = New String() {"Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье"}
    Private DeletedCount As Integer = 0
    Private FoundCount As Integer = 0
    Private UseDayFilter As Boolean = False
    Private FilesSize As Long
    Sub Main()
        Dim msg As String
        msg = "
***************************
iDeleter by Anton Vanichkin
***************************
"
        AddLog(msg)
        ' Инициализируем параметры обработки
        Dim TMass() As String = Split(Command.ToString, "/")
        Dim ParamList As New List(Of Params) From {
            New Params("pf ", False),
            New Params("sf ", False, "y"),
            New Params("wd ", False),
            New Params("minage ", False, 30),
            New Params("fm ", False),
            New Params("view", False),
            New Params("?", False),
            New Params("help", False)
        }
        For Each CMD As String In TMass
            For Each Param As Params In ParamList
                If (LCase(Mid(CMD, 1, Len(Param.Name))) = Param.Name) Then
                    Param.Value = Trim(Mid(CMD, Len(Param.Name) + 1))
                End If
            Next
        Next

        Dim FND As Params
        FND = ParamList.Find(Function(x) x.Name = "?")
        If FND.IsSet = True Then ShowHelpAndClose()

        FND = ParamList.Find(Function(x) x.Name = "help")
        If FND.IsSet = True Then ShowHelpAndClose()

        ' Проверяем наличие всех необходимых параметров
        For Each Result As Params In ParamList
            If Result.Reqired = True And Result.IsSet = False Then
                AddLog("Отсутствует обязательный параметр '" & RTrim(Result.Name) & "'")
                End
            End If
        Next

        ' Проверяем параметры на соответсвие типу и задаем переменные
        Dim Err As Boolean = False
        'Путь обработки
        FND = ParamList.Find(Function(x) x.Name = "pf ")
        If FND.IsSet = False Then
            WorkPath = AppDomain.CurrentDomain.BaseDirectory 'используем путь запуска приложения
        Else
            WorkPath = Replace(FND.Value, Chr(34), "")
            If IO.Directory.Exists(WorkPath) = False Then
                Err = True
                AddLog("Не найден каталог для обработки " & WorkPath)
            End If
        End If
        'Маска файлов
        FND = ParamList.Find(Function(x) x.Name = "fm ")
        If FND.IsSet = False Then
            FileMask = "*.*"
        Else
            FileMask = FND.Value
        End If
        'Обработка подкаталогов
        FND = ParamList.Find(Function(x) x.Name = "sf ")
        If LCase(FND.Value) = "n" Then SubFolders = False
        'Белые дни
        FND = ParamList.Find(Function(x) x.Name = "wd ")
        If FND.IsSet = True Then
            UseDayFilter = True
            TMass = Split(FND.Value, ",")
            Dim WDCount As Integer = 0
            ReDim WhiteDays(WDCount)
            For Each WD As Object In TMass
                If IsNumeric(WD) Then
                    If WD >= 1 And WD <= 7 Then
                        If WDCount <> 0 Then
                            ReDim Preserve WhiteDays(WDCount)
                        End If
                        WhiteDays(WDCount) = WD
                        WDCount = WDCount + 1
                    Else
                        Err = True
                        AddLog("День не может быть равен " & WD & ". Допустимый диапазон от 1 до 7")
                        Exit For
                    End If
                Else
                    Err = True
                    AddLog("День исключения указан неверно! " & WD)
                    Exit For
                End If
            Next
        End If
        'Минимальное время жизни
        FND = ParamList.Find(Function(x) x.Name = "minage ")
        If IsNumeric(FND.Value) Then MinAge = FND.Value
        'Режим просмотра списка
        FND = ParamList.Find(Function(x) x.Name = "view")
        If FND.IsSet Then
            ViewOnly = True
        End If
        If Err = True Then
            AddLog("Запустите утилиту без параметров либо с ключом /? или /help для получения справки")
            Console.ReadKey()
            End
        End If
        'Выводим результирующие параметры выполнения 
        msg = "
Заданные параметры
Путь для обработки: {0}
Режим просмотра: {1}
Маска файлов: {2}
Обработка подкаталогов: {3}
Исключать файлы младше {4}дн.
Исключение дней создания файла: {5}
***************************
"
        Dim DaysTX As String = ""
        If (UseDayFilter = True) Then
            msg = msg & "Исключение дней создания файла: "
            Dim zpt As String = ""
            For Each WD As Integer In WhiteDays
                DaysTX &= zpt & Days(WD - 1)
                zpt = ", "
            Next
        End If
        AddLog(String.Format(msg, WorkPath, ViewOnly.ToString, FileMask, SubFolders.ToString, MinAge, DaysTX))

        ' Всё готово для выполнения обработки
        ProcessFolder(WorkPath)
        Dim SizeFree As Long = FilesSize / 1048576
        msg = "
Найдено файлов, соответствующих фильтру: {0}
Размер файлов: {1}Mb
Удалено: {2}
"
        AddLog(String.Format(msg, FoundCount, SizeFree, DeletedCount))
    End Sub
    Private Sub ShowHelpAndClose()
        Dim msg As String = "
Утилита iDeleter предназначена для удаления файлов с датой создания не соответствующей заданному фильтру.
Параметры запуска: 
/pf <путь> - process folder, путь к папке, в которой будет производиться операция удаления.
    Если параметр пустой, обработка будет проводиться в каталоге с утилитой.
/fm <*.*> - file mask, маска отбора по типу файлов для удаления.
/sf [y|n] - subfolders, производить обработку в подкаталогах. По умолчанию - поиск в подкаталогах включен.
/wd <1,2...7> - white days, список дней недели, через запятую (1 - понедельник, 7 - воскресенье). Файлы созданные в эти дни удаляться не будут.
Например, если данный параметр равен 6, то файлы созданные в субботу будут сохранены.
/minage <цифра> - число дней. Файлы созданные менее, чем указано в этом параметре дней, удаляться не будут.
/view - только отображение информации по выполняемым операциям, без удаления файлов. Для отладки параметров запуска."
        AddLog(msg)
        Console.ReadKey()
        End
    End Sub
    Private Sub ProcessFolder(ByVal Path As String)
        ' Перебирваем файлы в данной папке
        Dim Files() As String
        Try
            Files = Directory.GetFiles(Path, FileMask)
        Catch ex As Exception
            AddLog("Ошибка открытия каталога " & Path)
            Exit Sub
        End Try
        For Each F As String In Files
            Dim Info As FileInfo = New FileInfo(F)
            Dim CreateDate As Date = Info.LastWriteTime
            Dim FileDay As Integer
            Dim Age As Long = DateDiff(DateInterval.Day, CreateDate, Now)
            FileDay = CreateDate.DayOfWeek
            If FileDay = 0 Then FileDay = 7
            If Age >= MinAge And InWhiteList(FileDay) = False Then
                AddLog(F & " " & CreateDate & " " & Days(FileDay - 1) & " возраст: " & Age)
                FoundCount = FoundCount + 1
                FilesSize = FilesSize + Info.Length
                If ViewOnly = False Then
                    'Удаление файла
                    Try
                        File.Delete(F)
                        DeletedCount += 1
                    Catch ex As Exception
                        AddLog("!!! Ошибка удаления " & F & ". " & ex.Message)
                    End Try
                End If
            End If
        Next

        ' Перебираем папки в данной папке 
        If SubFolders = False Then Exit Sub
        Dim Folders() As String = Directory.GetDirectories(Path)
        For Each F As String In Folders
            ProcessFolder(F)
        Next
    End Sub
    Private Function InWhiteList(ByVal Day As Integer) As Boolean
        For Each WD As Integer In WhiteDays
            If WD = Day Then
                Return True
            End If
        Next
        Return False
    End Function
    Private Sub AddLog(TX As String)
        Console.WriteLine(TX)
        Debug.Print(TX)
    End Sub
End Module
