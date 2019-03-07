#############################################################
##        Set aging recursively on Outlook folder          ##
##                     +++++++++++++++                     ##
##               Sergey Krasnov, 27.09.2017                ##
#############################################################

cls
Write-Output "=================================================================="

# Add Outlook .Net class
try
{
    Add-Type -assembly "Microsoft.Office.Interop.Outlook"
}
catch
{
    Write-Output "ERROR: Cannot add Microsoft.Office.Interop.Outlook assembly. Exit now."
    Exit
}


# Create COM-Object connection
try
{
    $app = New-Object -ComObject Outlook.Application
}
catch
{
    Write-Output "ERROR: Cannot Create COM-Object connection Outlook.Application. Exit now."
    Exit
}

# Add namespace
try
{
    $namespace = $app.GetNamespace("MAPI")

}
catch
{
    Write-Output "ERROR: Cannot Create MAPI connection. Exit now."
    Exit
}


# Set MAPI attributes we want to change in Outlook folders
# If you are planning GPO script deployment you only need $PR_AGING_DEFAULT
$PR_AGING_AGE_FOLDER = "http://schemas.microsoft.com/mapi/proptag/0x6857000B"
$PR_AGING_PERIOD = "http://schemas.microsoft.com/mapi/proptag/0x36EC0003"
$PR_AGING_GRANULARITY = "http://schemas.microsoft.com/mapi/proptag/0x36EE0003"
$PR_AGING_DELETE_ITEMS = "http://schemas.microsoft.com/mapi/proptag/0x6855000B"
$PR_AGING_FILE_NAME_AFTER9 = "http://schemas.microsoft.com/mapi/proptag/0x6859001E"
$PR_AGING_DEFAULT = "http://schemas.microsoft.com/mapi/proptag/0x685E0003"

# Accesibility MAPI attribute which is needed cause if we'll try to change public folder setting we'll get an exception
$PR_ACCESS = "http://schemas.microsoft.com/mapi/proptag/0x0FF40003"

# Set aging parameters
# $Granularity - sets what units is used for aging time
# 0 - months, 1 - weeks, 2 - days
#
# $Period - set aging time period. Elements that old is treated as expired.
# you can set 0-999
#
# $AgeFolder - allow folder aging
#
# $DeleteItems - allow expired items deletion
#
# $Default - should we use default settings
# 0 - None default setting is used
# 1 - Archive file default location is used only. Flags  are ticked только расположение файла архива берется по-умолчанию, стоят галочки 
# "Архивировать папку со следующими настройками" и "Перемещать старые элементы в папку по-умолчанию" 
# 3 - Все настройки установлены берутся по-умолчанию, стоит галочка 
# "Архивировать элементы папки с настройками по-умолчанию" 
$Granularity=2
$Period=6
$AgeFolder=$True
$Default=3
$DeleteItems=$false

# Проверяем валидность параметров
if ($Granularity -lt 0 -or $Granularity -gt 2) 
{
    Write-Output "ОШИБКА: Неверный параметр $Granularity. Выход из скрипта."
    Exit
}

if ($Period -lt 1 -or $Period -gt 999)
{
    Write-Output "ОШИБКА: Неверный параметр $Period. Выход из скрипта."
    Exit
}

if ($Default -lt 0 -or $Default -eq 2 -or $Default -gt 3)
{
    Write-Output "ОШИБКА: Неверный параметр $Default. Выход из скрипта."
    Exit
}


# Функция, которую будем вызывать последовательно на каждую папку
Function SetAgingOnFolder ($objFolder)
{
    # Проверки на допустимость входных параметров функции
	if ($objFolder -eq $null)
	{
		return $False
        Write-Output ("ОШИБКА: Неверный объект папки, переданный на вход функции (SetAgingOnFolder), имя папки:" + $objFolder.Name.ToString())
	}
	
	Try
	{
        # Получаем доступ к скрытому элементу папки, в котором хранятся настройки автоархивации
		$oStorage = $objFolder.GetStorage("IPC.MS.Outlook.AgingProperties", 2)
        # Интерфейс доступа к свойствам элемента
		$oPA = $oStorage.PropertyAccessor
		
        
		If ($AgeFolder -eq $false)
		{
			oPA.SetProperty PR_AGING_AGE_FOLDER, False
		}
		Else
		{
			# Устанавливаем настройки архивации папки, раскомментировать нужные, в нашем случае нужен только Default
			#$oPA.SetProperty($PR_AGING_AGE_FOLDER, $True)
			#$oPA.SetProperty($PR_AGING_GRANULARITY, $Granularity)
			#$oPA.SetProperty($PR_AGING_DELETE_ITEMS, $DeleteItems)
			#$oPA.SetProperty($PR_AGING_PERIOD, $Period)
			$oPA.SetProperty($PR_AGING_DEFAULT, $Default)
            # В том числе и имя файла, если задано
			if ($FileName -ne $null)
			{
				 #$oPA.SetProperty($PR_AGING_FILE_NAME_AFTER9, $FileName)
			}

		}

		# Сохраняем изменения в скрытом элементе папки и возвращаем успешное выполнение функции
		$oStorage.Save()
        Write-Output ("УСПЕХ: Установка аттрибутов автоархивации на папку >>> " + $objFolder.Name.ToString())
	}
	
    # Если ошибка, то пишем в консоль и возвращаем неудачное выполнение функции
	Catch
	{
        Write-Output ("ОШИБКА: Что-то пошло не так при установке параметров в функции (SetAgingOnFolder)... " + $objFolder.Name.ToString())
		Write-Output $_
	}
}

# Функция, которая будет проходить по всем подпапкам текущей папки и вызывать для каждой установку
# параметров автоархивации
Function RecursiveRoutine ($objParent)
{
#    SetAgingOnFolder ($objFolder)
    $objChildren = $objParent
    foreach ($objChild in $objChildren)
    {
        SetAgingOnFolder ($objChild)
        RecursiveRoutine ($objChild.folders)
    }
}

# Основной код
###############
Write-Output "Ищем настроенный профиль..."

# Проверяем, есть ли настроенный профиль Outlook
try
{
    if ($app.DefaultProfileName -eq $null)
    {
        Write-Output "ОШИБКА: Не настроен профиль по-молчанию. Выход из скрипта."
        Exit
    }
}
catch
{
    Write-Output "ОШИБКА: Не могу получить имя профиля по-умолчанию. Выход из скрипта."
    Exit
}


# Проверяем, есть ли настроенные аккаунты электронной почты в Outlook
try
{
    $accCount = ($app.Explorers.Session.Accounts | Measure-Object).count
    if ($accCount -eq 0)
    {
        Write-Output "ОШИБКА: Не сконфигурировано ни одного аккаунта электронной почты. Выход из скрипта."
        Exit
    }
}
catch
{
    Write-Output "ОШИБКА: Не могу получить количество настроенных аккаунтов электронной почты. Выход из скрипта."
    Exit
}

# Получаем список объектов папок
try
{
    $objAllfolders = $namespace.Session.Folders
}
catch
{
    Write-Output "ОШИБКА: Не могу получить папки аккаунта в Outlook. Выход из скрипта."
    Exit
}


# Каждую корневую папку, кроме "общедоступных папок", пропускаем по рекурсивной функции
ForEach ($objFolder in $objAllfolders) 
{
    Write-Output " "
    Write-Output "====================================================="
    Write-Output ("Заходим в корень папки: " + $objFolder.Name)

    try
    {
        $accessLevel = $objFolder.PropertyAccessor.GetProperty($PR_ACCESS)
    }
    catch
    {
        Write-Output "ОШИБКА: Не могу получить аттрибут PR_ACCESS, пропускаю папку"
        Continue
    }

    if ($accessLevel -eq 63)
    {
        RecursiveRoutine $objFolder.Folders
    }
    else
    {
        Write-Output "Системная папка, пропускаем..."
        Continue
    }
}
