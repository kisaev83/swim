
def instruction_menu(button):
    if button == 'Главное меню':
           text = "*Справка. Главное меню.*\n\n" \
                  "При нажатии кнопки /start или при возврате из турнира Вы попадаете в главное меню и " \
                  "Вам доступны следующие кнопки:\n\n" \
                  "*Турниры в работе* - в этом подменю будут доступны все созданные Вами турниры с датой проведения большей " \
                  "или равной сегодняшней дате. Кнопка доступна при создании одного или более турниров.\n\n" \
                  "*Архивные турниры* - такое же подменю, как и 'Турниры в работе', только в нём доступны уже прошедшие Ваши турниры.\n\n" \
                  "*Присоединиться к турниру* - данный пункт меню доступен для помощников в турнире. Администратор турнира из " \
                  "меню турнира сможет пригласить в этот турнир помощников по ссылке-приглашению. " \
                  "Инструкция по этому меню будет доступна как для помощников, так и для администратора турнира при выборе " \
                  "кнопки 'Добавить помощников в турнир' из меню турнира.\n\n" \
                  "*Создать новый турнир* - создание нового турнира. Следуйте указаниям бота.\n\n" \
                  "*Справка о боте* - эта справка."
    elif button == 'Меню турнира 1':
           text = "*Справка. Меню турнира (1).*\n\n" \
                  "В зависимости от заполнения турнира " \
                  "информацией, пункты меню выбранного турнира будут меняться.\n" \
                  "Так как бот формирует текстовые сообщения с кнопками и они не меняются без действий пользователя, " \
                  "и если при каких-либо обстоятельствах меню не поменялось, необходимо выйти в главное меню и " \
                  "войти в меню турнира заново.\n\n" \
                  "При формировании ботом любого файла, к файлу будет добавлено описание турнира. Это " \
                  "обычное сообщение телеграм, которое Вы сможете переслать в свою группу или сохранить файл на устройство " \
                  "и отправить его любым удобным способом.\n\n" \
                  "После создания нового турнира Вам будут доступны 4 пункта меню, а именно:\n" \
                  "*Сформировать техническую заявку.* Техническая заявка - это первое с чего начинается каждый турнир, " \
                  "не считая положения. В ответ вам придёт сообщение с файлом технической заявки в формате Excel. В нём уже " \
                  "проставлены выбранные Вами дистанции. Этот файл необходимо разослать всем вашим участникам. " \
                  "Участники его заполняют и отправляют Вам, указанным Вами способом. Файл защищен от изменений, доступны " \
                  "только зеленые поля для заполнения. Инструкция по заполнению заявки находится в одной из ячеек " \
                  "в этом файле.\n\n" \
                  "*Изменить/удалить турнир* - следуйте инструкциям бота.\n\n" \
                  "*Добавить помощников в турнир* - следуйте инструкциям бота.\n\n" \
                  "*Загрузить техническую заявку от команды* - с помощью этой кнопки Вы будете загружать " \
                  "заполненные участниками технические заявки. После приглашения ботом об отправке файла, приложите в свое сообщение, " \
                  "как документ, файл с заявкой с расширением xlsx.\nОбратите внимание! " \
                  "Ботом обрабатываются только заявки по форме, файл которой был сформирован на этапе 'Сформировать заявку'.\n" \
                  "*Важное примечание!* Работа с заявкой от одной команды происходит через один файл от этой команды! " \
                  "При любых изменениях в заявке - название команды должно оставаться таким же!!! Если в названии команды " \
                  "будет изменен хотя бы один символ - это будет уже другая команда и в базе данных продублируются те " \
                  "участники, которые ранее уже были добавлены. Если нужно удалить из турнира всю команду, необходимо " \
                  "в файле заявки этой команды очистить ячейки с участниками(можно только фамилии) " \
                  "(именно очистить, а не удалить. Удалить строки файл запретит), и главное оставить " \
                  "название команды таким же, которое уже было добавлено и загрузить этот файл в бота.\nТакже поступайте " \
                  "с участниками, которые снимаются с турнира до формирования стартового протокола - очистите данные участника, " \
                  "при этом не изменяя остальных, " \
                  "сохраните файл и загрузите его вновь.\n" \
                  "То же самое происходит с заявками, где название команд совпадает, а участники в них разные. " \
                  "Бот воспримет название команды, как одну, и удалит из турнира участников из первой загруженной заявки и добавит " \
                  "из второй. Для этого либо объедините участников в один файл заявки, либо переименуйте название команды " \
                  "(добавьте символ в название) в одной из заявок.\n\n" \
                  "После успешной обработки заявки, бот напишет сообщение с информацией о загруженной заявке."
    elif button == 'Меню турнира 2':
           text = "*Справка. Меню турнира (2).*\n\n" \
                  "После добавления первой заявки от команды в меню турнира добавятся следующие кнопки:\n\n" \
                  "*Заявочный протокол* - в ответ сформируется файл в формате pdf - заявочный протокол по дистанциям " \
                  "с сортировкой всех участников по заявочному времени. Этот файл может быть отправлен участникам для " \
                  "сверки, когда приём заявок окончен.\n\n" \
                  "*Стартовый протокол* - важный протокол турнира. После нажатия этой кнопки вам будет предложено " \
                  "сформировать стартовый протокол по Вашим параметрам.\n" \
                  "1. Определить порядок формирования. У вас есть 2 варианта:\n" \
                  "'Все участники по времени' - объединит всех участников турнира из всех возрастных " \
                  "категорий и сформирует стартовый протокол с сортировкой по заявочному времени.\n" \
                  "'Сгруппировать участников по годам' - вы самостоятельно группируете участников в заплывы по возрасту. " \
                  "И они будут чередоваться девушки-юноши.\n" \
                  "2. 'Кто первые начинают турнир?' - девушки или юноши.\n" \
                  "3. 'Какое минимальное количество участников в последнем заплыве группы?' - описание лучше подойдет " \
                  "примером: Допустим, у Вас 6 дорожек в бассейне и 19 юношей в возрастной категории, то в обычной ситуации " \
                  "у вас получится 3 полных заплыва по 6 человек и последний 4-й заплыв будет состоять из одного участника. " \
                  "Этим параметром вы выбираете минимальное количество участников в последнем заплыве. " \
                  "Если, допустим, Вы выбрали 3, то стартовый протокол сформируется следующим образом: " \
                  "2 первых заплыва будут по 6 человек, в предпоследнем, 3-м, заплыве " \
                  "будет 4 участника, а в последнем, 4-м, - 3 участника.\n" \
                  "*Важное примечание!* Каждый новый сформированный стартовый протокол - будет непохож на другой. Т.к. " \
                  "используется функция случайности. Участники с одинаковым заявочным временем случайным образом " \
                  "формируются в заплывы. В одном протоколе Иванов может попасть в 3-й заплыв на 6 дорожку, а в " \
                  "другом протоколе в 4-й заплыв на 3 дорожку.\n\n" \
                  "После выбора всех параметров бот сформирует несколько файлов:\n" \
                  "1. Собственно сам стартовый протокол в формате pdf, который можно отправлять участникам и выводить " \
                  "на печать.\n" \
                  "2. Файл Дорожки.pdf - файл для вывода на печать. Для судей на секундомерах. У каждого судьи на дорожке " \
                  "только участники плывущие по его дорожке. Необходим при ручном хронометраже и фиксирования результатов в заплывах.\n" \
                  "3. Файлы ОРГАНИЗАТОР - формат Excel. В эти файлы вносятся результаты заплывов и во время проведения " \
                  "турнира загружаются в бот. Сохраните эти файлы в удобном месте на своем устройстве, для дальнейшей " \
                  "загрузки в бот.\n\n" \
                  "*Списки участников* - в этом пункте меню есть 2 кнопки подменю:\n" \
                  "'Список участников по категориям' - формирует pdf файл по форме: Дистанция, пол, возрастная категория. " \
                  "Сортировка участников по фамилии. Этот файл также может быть отправлен для сверки участникам на " \
                  "Ваше усмотрение.\n" \
                  "'Общее количество участников' - формирует pdf файл с общим количеством участников по командам и по " \
                  "году рождения.\n\n" \
                  "На данном этапе ведения турнира из меню уберется кнопка 'Сформировать техническую заявку'" \

    elif button == 'Меню турнира 3':
           text = "*Справка. Меню турнира (3).*\n\n" \
                  "В день проведения турнира добавится кнопка *'Загрузить результаты'*\n" \
                  "Вы или ваш помощник заполняете результатами файл Excel, присланный ботом " \
                  "при формировании стартового протокола с названием ОРГАНИЗАТОР...\n" \
                  "После приглашения ботом загрузить файл, отправьте файл с расширением xlsx. После обработки " \
                  "файла, бот пришлет сообщение об удачной обработке результатов.\n" \
                  "Примечание. Файл можно загружать сколько угодно раз и в любое время, " \
                  "хоть после каждого заплыва, хоть после добавления каждого результата.\n" \
                  "*Важное примечание.* Не забывайте сохранять на устройстве этот файл " \
                  "с внесенными изменениями перед отправкой в бот.\n\n" \
                  "После первой загрузки результатов, меню турнира вновь изменится и добавятся следующие пункты:\n\n" \
                  "*Текущие результаты* - результаты по категориям и дистанциям. Будет доступно " \
                  "подменю с выбором категории, дистанции и пола. После выбора нужной дистанции, сформируется " \
                  "pdf файл с результатами. Его можно использовать для награждения, а также для информирования " \
                  "участников турнира, отправив им этот файл.\n" \
                  "*Обратите внимание* на текст сообщения от бота после нажатия 'Текущие результаты'. По мере " \
                  "заполнения базы данных результатами в тексте сообщения будут появляться строки 'Результаты " \
                  "добавлены не полностью'. Это означает, что количество заявившихся участников на эту дистанцию " \
                  "не совпадает с количеством внесенных результатов, т.е. либо не все участники категории " \
                  "проплыли эту дистанцию, либо не все результаты " \
                  "записались в базу. Бот сформирует файл, но там где не хватает результатов, в этой категории " \
                  "не будут указаны занятые участниками места - это промежуточные результаты. Они обновятся, когда все " \
                  "участники категории проплывут ту или иную дистанцию и будут загружен обновленный файл с резальтатами.\n\n" \
                  "*Итоговый протокол* - на данном этапе ведения турнира уже доступно формирование итогового протокола. " \
                  "Протокол формируется в формате pdf. Следуйте инструкциям бота.\n\n" \
                  "*Медалисты* - формируется файл pdf со всеми медалистами турнира. Файл со статистикой. Используется, как " \
                  "бонус, для отправки участникам по окончании турнира.\n\n" \
                  "*Многоборье по очкам FINA2022* - выберите дистанции для подсчета многоборья. Формируется файл pdf.\n\n" \
                  "На данном этапе ведения турнира уберутся пункты меню: 'Заявочный протокол', 'Стартовый протокол', " \
                  "'Загрузить техническую заявку от команды'" \

    return text



# "*Загрузить результаты* - кнопка для загрузки результатов заплывов во время проведения турнира. " \
