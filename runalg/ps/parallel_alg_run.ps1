#количество потоков параллельных
$MaxCountCmd    = 2

#куда генерировать временные файлы-bat 
$TmpPath       = 'C:\GalaktikaCorp\_runalg\tmp'

$list_algname = @(
('000100000000000Ah','Распределение заказов по грузовикам'),
('0001000000000048h','Расчет "Потребность в материалах" (Автомобили серия)'),
('0001000000000049h','Расчет "Потребность в материалах" (Запчасти)'),
('000100000000004Ah','Расчет "Потребность в материалах" (Кооперация)'),
('0001000000000047h','Расчет "Потребность в материалах" (Сборочные комплекты)'),
('0001000000000043h','Расчет "Внешняя поставка деталей" (Автомобили серия)'),
('0001000000000046h','Расчет "Внешняя поставка деталей" (Запчасти)'),
('0001000000000045h','Расчет "Внешняя поставка деталей" (Кооперация)'),
('0001000000000044h','Расчет "Внешняя поставка деталей" (Сборочные комплекты)'),
('000100000000003Fh','Расчет "Внутреннее производство деталей" (Автомобили серия)'),
('0001000000000040h','Расчет "Внутреннее производство деталей" (Запчасти)'),
('0001000000000041h','Расчет "Внутреннее производство деталей" (Кооперация)'),
('0001000000000042h','Расчет "Внутреннее производство деталей" (Сборочные комплекты)')

)
#Строка запуска галактики 
$GalExe = 'start C:\GalaktikaCorp\GAL91\exe\galnet.exe /c:"C:\GalaktikaCorp\GAL91\Start\test.cfg" '


#зачищаем все временные папки из $TmpPath
Write-Host 'Удаляем все папки во временной папке' -ForegroundColor Green
foreach($onefolderToDel in $(Get-ChildItem -Path $TmpPath -Directory))
{
  Remove-Item -Path $($onefolderToDel.FullName+'\RunGal.bat') -Recurse -Force
  Remove-Item -Path $onefolderToDel.FullName -Recurse -Force
}





# генерируем батник запуска галки
function generateFileExport([string]$fnFolderTmp ,[String]$value)
{
 
 New-Item -ItemType "directory" -Path $fnFolderTmp 
 cd $fnFolderTmp
 
 #$value = $value | ConvertTo-Encoding windows-1251 cp866

 Add-Content -Path "$fnFolderTmp\RunGal.bat" -value $value -Encoding Oem  #формат даты 31012021  
}

#создаем временные папки откуда галку будем запускать
foreach ($name in $list_algname )
{

 write-host $name[0]
 $foldertmp = $TmpPath + "\"+$name[1]

 $foldertmp = $foldertmp.Replace('"','')  
 $foldertmp = $foldertmp.Replace(' ','_')  
 $nrec = $name[0]
 generateFileExport -fnFolderTmp  $foldertmp -value $("$GalExe/galaxy.nowrun=M_MNPLAN::RUNALG($nrec)")
 
}
 
 

#
#Cобственно сам запуск. Идем по созданным ранее папкам и в каждой запускаем батник
#
foreach($onefolderToExport in $(Get-ChildItem -Path $TmpPath -Directory ))
{

  do
  {
    $result = Get-Process -name atlexec -ErrorAction SilentlyContinue
         
    if ($result -eq $null) {
    $countcmd = 0
    }
    else {
    $countcmd = $result.Count
    }
    Start-Sleep 1
  } while ($countcmd -gt $MaxCountCmd)

  write-host $onefolderToExport.Name  -ForegroundColor Green

  $bat = $onefolderToExport.FullName + '\RunGal.bat' 
  cd $onefolderToExport.FullName
  
  if ($(Test-Path $bat) -eq $true) {
      
      Start-Process -FilePath $bat 
   }  
}
