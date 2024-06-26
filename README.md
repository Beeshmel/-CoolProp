# Инструкция для подключения библиотеки CoolProp Python и установки надстройки в Excel
CoolProp является открытым программным обеспечением (Open Source) и разработана сообществом разработчиков под лицензией MIT. Это означает, что программу CoolProp можно использовать бесплатно как в коммерческих, так и в некоммерческих целях без каких-либо ограничений.\
CoolProp не имеет конкретного коммерческого производителя, так как она разрабатывается и поддерживается открытым сообществом. Разработка программы осуществляется на платформе GitHub, где любой желающий может принять участие в развитии проекта, внести свой вклад и предложить улучшения.\
Таким образом, благодаря своему открытому и бесплатному характеру CoolProp стала широко используемым инструментом в области теплофизики, технического расчета и других инженерных областей, позволяя пользователям проводить точные и эффективные расчеты термодинамических свойств веществ и смесей.\
Программа CoolProp основана на уравнениях состояния и моделях, позволяющих с высокой точностью и эффективностью рассчитывать термодинамические свойства с различными входными параметрами, такими как давление, температура и состав смеси. CoolProp поддерживает большое количество веществ и смесей, включая углеводороды, фреоны, пары и многое другое.\
Основные возможности CoolProp для Python и Excel включают:
1. Расчет термодинамических свойств: CoolProp позволяет легко и быстро рассчитывать термодинамические свойства веществ и смесей для широкого диапазона условий.
2. Совместимость с Python: CoolProp является библиотекой на языке Python, что обеспечивает простоту интеграции и использования в различных прикладных задачах.
3. Интерфейс для Excel: CoolProp также имеет интерфейс для Microsoft Excel, что позволяет проводить расчеты термодинамических свойств непосредственно в таблицах Excel.
4. Расчеты теплообмена: с помощью CoolProp можно проводить расчеты теплообмена и тепловых процессов, используя точные данные о термодинамических свойствах веществ.
5. Поддержка различных единиц измерения: CoolProp поддерживает работу с различными системами единиц измерения, обеспечивая удобство использования для пользователей.

Таким образом, CoolProp представляет собой мощный инструмент для расчета теплофизических свойств веществ и смесей в различных приложениях на Python и Excel, обеспечивая точность, эффективность и удобство использования для пользователей. Благодаря возможности бесплатного использования и открытому коду CoolProp является популярным выбором среди специалистов в области теплофизики, химии, инженерии и других смежных областей.


## Подключение библиотеки CoolProp в Python
Скачайте последнюю версию библиотеки запустив в Командной строке: pip install CoolProp\
Добавить библиотеку CoolProp в проект можно в Настройках -Python interpreter.\
Если на компьютере установлен REFPROP, то CoolProp позволяет подключить и библиотеку данной программы.
 
``` from CoolProp.CoolProp import PropsSI
import CoolProp.CoolProp as CP

composition = 'METHANE[0.90]&ETHANE[0.07]&PROPANE[0.03]'
Tin = 300  # K
Pin = 9.61 * 10 ** 6  # Pa
H = PropsSI("H" , "T" , Tin , "P" , Pin , composition) 

CP.set_config_string(CP.ALTERNATIVE_REFPROP_LIBRARY_PATH , 'C:\Program Files (x86)\REFPROP\REFPRP64.DLL')
CP.set_config_bool(CP.REFPROP_USE_GERG , True)

composition_REF = 'REFPROP::METHANE[0.90]&ETHANE[0.07]&PROPANE[0.03]'
H = PropsSI("H" , "T" , Tin , "P" , Pin , composition_REF)
```

## Подключение библиотеки CoolProp в Excel
Скачивание программы для работы с CoolProp необязательно, для этого достаточно использовать дополнительные плагины. \
Скачать программу можно на сайте: \
https://sourceforge.net/projects/coolprop/files/CoolProp/6.6.0/Installers/Windows/

- Скачайте файлы CoolProp.dll и CoolProp_x64.dll и поместите их в папку C:\CoolProp*.
- Скачайте файл CoolProp.xlam и поместите его в удобное место
- Установите надстройку CoolProp из меню Параметры Excel, Управление надстройками
   1. Откройте Excel
   2. Перейдите в меню Файл→Параметры→Дополнительные модули
   3. В нижней части панели выберите Управление: Надстройки Excel, затем нажмите кнопку Перейти..
   4. Нажмите кнопку "Обзор" на панели "Менеджер надстроек".
   5. Перейдите к файлу CoolProp.xlam, который вы скопировали выше в шаге (b).
   6. Убедитесь, что надстройка CoolProp выбрана (флажок установлен), и закройте диспетчер надстроек.
- Откройте файл TextExcel.xlsx (этот файл) и попробуйте повторно оценить одну из ячеек с формулами на вкладке Sample Calcs. Все формулы CoolProp должны работать.

