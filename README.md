# deployOffice 2021
deployOffice 2021

Скрипт предназначен для разворачивания __MS Office__ 2021 или др. по принципу __click to run.__
	
У Вас должен быть __предварительно сохранен локально__ дистрибутив MS Office с сети CDN разрядности 32 и 64.
Путь до дистрибутива указывается в config.ini, ...\deploy2021 версии по разрядности организуются соответсвенно \32 и \64.
В корень deploy2021 требуется положить файл разворачивания setup.exe
  
Особенность скрипта в том, что для установки ПО нет необходимости подключения к целевому ПК и запуска скрипта локально.
  
Для этого администратор запускает задачу со своей машины, указывая имя целевого ПК.
Да, для выполнения необходимы права доступа локального администратора и на удаленной машине включен admin$.
	
Логин локального администратора указывается в config.ini , пароль запрашивается в процессе выполнения.

Скрипт автоматически запрашивает разрядность целевого ПК, по выбору администатора __компонентов пакета__ - подготавливает XML файл,
и передает задачу на исполнение целевому ПК на исполнение. Далее целевой ПК самостоятельно выполнит установку. Для пользователя выполнение установки скрыто.
