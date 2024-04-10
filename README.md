# -CoolProp
#Инструкция для подключения библиотеки CoolProp Python и установки надстройки в Excel

Подключение библиотеки CoolProp в Python
Скачайте последнюю версию библиотеки запустив в Командной строке: pip install CoolProp
Добавить библиотеку CoolProp в проект можно в Настройках -Python interpreter.
Если на компьютере установлен REFPROP, то CoolProp позволяет подключить и библиотеку данной программы. 
 
from CoolProp.CoolProp import PropsSI
import CoolProp.CoolProp as CP
composition = 'METHANE[0.90]&ETHANE[0.07]&PROPANE[0.03]'
Tin = 300  # K
Pin = 9.61 * 10 ** 6  # Pa
H = PropsSI("H" , "T" , Tin , "P" , Pin , composition)

CP.set_config_string(CP.ALTERNATIVE_REFPROP_LIBRARY_PATH , 'C:\Program Files (x86)\REFPROP\REFPRP64.DLL')
CP.set_config_bool(CP.REFPROP_USE_GERG , True)
composition_REF = 'REFPROP::METHANE[0.90]&ETHANE[0.07]&PROPANE[0.03]'
H = PropsSI("H" , "T" , Tin , "P" , Pin , composition_REF)

#Подключение библиотеки CoolProp в Excel
Скачивание программы для работы с CoolProp необязательно, для этого достаточно использовать дополнительные плагины. 
Скачать программу можно на сайте: 
https://sourceforge.net/projects/coolprop/files/CoolProp/6.6.0/Installers/Windows/
1.	Скачайте файлы CoolProp.dll и CoolProp_x64.dll и поместите их в папку C:\CoolProp*.
2.	Скачайте файл CoolProp.xlam и поместите его в удобное место
3.	Установите надстройку CoolProp из меню Параметры Excel, Управление надстройками
  a.	Откройте Excel
  b.	Перейдите в меню Файл→Параметры→Дополнительные модули
  c.	В нижней части панели выберите Управление: Надстройки Excel, затем нажмите кнопку Перейти....
  d.	Нажмите кнопку "Обзор" на панели "Менеджер надстроек".
  e.	Перейдите к файлу CoolProp.xlam, который вы скопировали выше в шаге (b).
  f.	Убедитесь, что надстройка CoolProp выбрана (флажок установлен), и закройте диспетчер надстроек.
4.	Откройте файл TextExcel.xlsx (этот файл) и попробуйте повторно оценить одну из ячеек с формулами на вкладке Sample Calcs. Все формулы CoolProp должны работать.

