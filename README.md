# deployOffice 2021

  	Скрипт предназначен для разворачивания MS Office 2021 или др. по принципу click to run.
  У Вас должен быть подготовлен и скачан предварительно с сети CDN дистрибутив MS Office разрядности 32 и 64.
  Путь до дистрибутива указывается в config.ini
  Особенность скрипта в том, что для установки ПО нет необходимости подключения к целевому ПК и запуска скрипта локально.
	Для этого администратор запускает задачу со своей машины, указывая имя целевого ПК.
	Да, для выполнения необходимы права доступа локального администратора и на удаленной машине включен admin$.
Логин локального администратора указывается в config.ini , пароль запрашивается в процессе выполнения.
	Скрипт автоматически запрашивает разрядность целевого ПК, по выбору администатора компонентов пакета подготавливает XML файл - задачу для деплоя,
и передает на исполнение целевому ПК. Для пользователя выполнение работы скрыто.
