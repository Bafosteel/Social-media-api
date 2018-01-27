import requests
import json
from tkinter import *
from tkinter.filedialog import askopenfilename
import openpyxl
# Token = '##############'

class Gui(Toplevel):

    def __init__(self, parent, title="Обработка файлов"):
        Toplevel.__init__(self, parent)
        parent.geometry("250x250+100+150")
        if title:
            self.title(title)
        parent.withdraw()
        self.parent = parent
        self.result = None
        dialog = Frame(self)
        self.initial_focus = self.dialog(dialog)
        self.protocol("WM_DELETE_WINDOW", self.on_exit)
        dialog.pack()

    def on_exit(self):
        self.quit()

    def search_folder_for_excel_file(self):
        path_to = askopenfilename()
        print(path_to)
        self.text_1.delete(0, END)
        self.text_1.insert(END, path_to)

    def get_video(self,path_to_file,name_of_video):
        Token = str(self.text_3.get())
        ids = []
        print('4pok')
        for i in range(0, len(path_to_file)):
            try:
                vid = {'video': (
                str(path_to_file[i]), open(r''+str(path_to_file[i]), 'rb'))}
                group_id_1 = str(self.text_2.get())
                method = 'https://api.vk.com/method/video.save?'
                data = dict(access_token=Token, gid=group_id_1,name=name_of_video[i])
                response = requests.post(method, data)
                result = json.loads(response.text)

                upload_url = result['response']['upload_url']
                response = requests.post(upload_url, files=vid)
                result = json.loads(response.text)

                ids.append('video-' + str(self.text_2.get()) + '_' + str(result['video_id']))
            except:
                print("This is something else")
                ids.append(str(path_to_file[i]))
        return ids

    def get_image_id(self,path_to_file):
        Token = str(self.text_3.get())
        # Получаем id изображения
        group_id_1 = str(self.text_2.get())
        id_s=[]
        for path in path_to_file:
            try:
                # путь к вашему изображению
                img = {'photo': (str(path), open(r''+str(path), 'rb'))}

                # Получаем ссылку для загрузки изображений
                method_url = 'https://api.vk.com/method/photos.getWallUploadServer?'
                data = dict(access_token=Token, gid=group_id_1)
                response = requests.post(method_url, data)
                result = json.loads(response.text)
                upload_url = result['response']['upload_url']

                # Загружаем изображение на url
                response = requests.post(upload_url, files=img)
                result = json.loads(response.text)

                # Сохраняем фото на сервере и получаем id
                method_url = 'https://api.vk.com/method/photos.saveWallPhoto?'
                data = dict(access_token=Token, gid=group_id_1, photo=result['photo'], hash=result['hash'],
                        server=result['server'])
                response = requests.post(method_url, data)
                result = json.loads(response.text)['response'][0]['id']
                id_s.append(result)
            except:
                print('This is href')
                id_s.append(str(path))

        return id_s

    def get_columns_data(self):
        # Получим столбцы из эксель файла
        wb = openpyxl.load_workbook(str(self.text_1.get()))
        sheet = wb.worksheets[0]
        private_post = []
        message = []
        path_to_file = []
        link_button = []
        link_title = []
        link_image = []
        post_with_button = []
        type_of_attachment = []
        name_of_video = []
        for i in range(1, sheet.max_row):
            if ((sheet.cell(row=i, column=1).value) == None):
                max_row = i - 1
                break
            else:
                max_row = sheet.max_row

        for i in range(2, max_row + 1):
            private_post.append(sheet.cell(row=i, column=1).value)
            post_with_button.append(sheet.cell(row=i, column=2).value)
            message.append(sheet.cell(row=i, column=3).value)
            type_of_attachment.append(sheet.cell(row=i, column=4).value)
            name_of_video.append(sheet.cell(row=i, column=5).value)
            path_to_file.append(str(sheet.cell(row=i, column=6).value).replace('\\','/'))
            link_button.append(sheet.cell(row=i, column=7).value)
            link_title.append(sheet.cell(row=i, column=8).value)
            link_image.append(sheet.cell(row=i, column=9).value)
        return private_post,post_with_button, message,type_of_attachment, \
               path_to_file,link_button,link_title,link_image,name_of_video

    def refill_table(self,posts):
        import ctypes
        wb = openpyxl.load_workbook(str(self.text_1.get()))
        sheet = wb.worksheets[0]
        for i in range(0, len(posts)):
            print(posts[i])
            sheet.cell(row=i+2, column=9).value = str(posts[i])
        wb.save(str(self.text_1.get()))
        message = 'Готово!'
        ctypes.windll.user32.MessageBoxW(0, message, 'Данные о каналах', 0)
        print('ok')
        return {}

    def start(self):
        Token = str(self.text_3.get())

        private_post,post_with_button, message,type_of_attachment,\
        path_to_file,link_button,link_title,link_image,name_of_video = self.get_columns_data()
        result = self.get_image_id(path_to_file)
        video = self.get_video(path_to_file,name_of_video)
        posts = []
        group_id = "-"+str(self.text_2.get())

        buttons = {'Запустить':'app_join','Играть':'app_game_join','Перейти':'open_url','Открыть':'open',
                   'Подробнее':'more','Позвонить':'call','Забронировать':'book','Записаться':'enroll',
                   'Зарегистрироваться':'register','Купить':'buy','Купить билет':'but_ticket','Заказать':'order',
                   'Установить':'install','Связаться':'contact','Заполнить':'fill','Подписаться':'join_public',
                   'Я пойду':'join_event','Вступить':'join','Связаться_2':'im','Написать':'im2'}

        # Создаем посты
        for i in range(0,len(private_post)):
            if ('photo' in type_of_attachment[i]):
                print('photo')
                if ("private" not in private_post[i]):
                    print("public")
                    # Создаем пост
                    r = requests.post('https://api.vk.com/method/wall.post',
                                    params={'owner_id': group_id, 'access_token': Token, 'from_group': 1,
                                              'message': str(message[i]),
                                              'attachments': str(result[i])})
                    response = r.json()
                    posts.append('https://vk.com/wall-'+str(group_id)[1:]+'_'+str(response['response']['post_id']))
                else:
                    print("private")
                    if ('yes' in post_with_button[i]):
                        # Создаем скрытый пост с кнопкой и с фото
                        r = requests.post('https://api.vk.com/method/wall.postAdsStealth',
                                        params={'owner_id': group_id, 'access_token': Token, 'from_group': 1,
                                                'message': str(message[i]),
                                                'link_button': str(buttons[link_button[i]]),
                                                'link_image': str(link_image[i]),
                                                'link_title': str(link_title[i]),
                                                'attachments': str(result[i])})
                        response = r.json()
                        print(response)
                        posts.append('https://vk.com/wall-' + str(group_id)[1:]+'_'+ str(response['response']['post_id']))
                    else:
                        # Создаем скрытый пост без кнопки и с фото
                        r = requests.post('https://api.vk.com/method/wall.postAdsStealth',
                                          params={'owner_id': group_id, 'access_token': Token, 'from_group': 1,
                                                  'message': str(message[i]),
                                                  'attachments': str(result[i])})
                        response = r.json()
                        print(response)
                        posts.append(
                            'https://vk.com/wall-' + str(group_id)[1:] + '_' + str(response['response']['post_id']))
            else:
                print('video')
                if ("private" not in private_post[i]):
                    print("public")
                    # Создаем пост
                    r = requests.post('https://api.vk.com/method/wall.post',
                                    params={'owner_id': group_id, 'access_token': Token, 'from_group': 1,
                                              'message': str(message[i]),
                                              'attachments': str(video[i])})
                    response = r.json()
                    posts.append('https://vk.com/wall-'+str(group_id)[1:]+'_'+str(response['response']['post_id']))
                else:
                    print("private")
                    if ('yes' in post_with_button[i]):
                        # Создаем скрытый пост с кнопкой и с видеозаписью
                        r = requests.post('https://api.vk.com/method/wall.postAdsStealth',
                                        params={'owner_id': group_id, 'access_token': Token, 'from_group': 1,
                                                'message': str(message[i]),
                                                'link_button': str(link_button[i]),
                                                'link_image': str(link_image[i]),
                                                'link_title': str(link_title[i]),
                                                'attachments': str(video[i])})
                        response = r.json()
                        print(response)
                        posts.append('https://vk.com/wall-' + str(group_id)[1:]+'_'+ str(response['response']['post_id']))
                    else:
                        # Создаем скрытый пост без кноки, но с видеозаписью
                        r = requests.post('https://api.vk.com/method/wall.postAdsStealth',
                                          params={'owner_id': group_id, 'access_token': Token, 'from_group': 1,
                                                  'message': str(message[i]),
                                                  'attachments': str(video[i])})
                        response = r.json()
                        print(response)
                        posts.append(
                            'https://vk.com/wall-' + str(group_id)[1:] + '_' + str(response['response']['post_id']))

        return self.refill_table(posts)

    def dialog(self, parent):
        self.parent = parent

        # Created main elements
        self.label_1 = Label(parent, text="Укажите путь, по которому лежит основной Excel файл")
        self.text_1 = Entry(parent, width=50)
        self.but_1 = Button(parent, text="Указать", command=self.search_folder_for_excel_file)

        self.label_2 = Label(parent, text="Укажите id сообщества")
        self.text_2 = Entry(parent)

        self.label_3 = Label(parent, text="Укажите Токен")
        self.text_3 = Entry(parent)

        self.label_1.pack()
        self.text_1.pack()
        self.but_1.pack()

        self.label_2.pack()
        self.text_2.pack()

        self.label_3.pack()
        self.text_3.pack()

        # start button
        self.but_start = Button(parent, text="Выполнить",command=self.start)
        self.but_start.pack()



if __name__ == "__main__":
        root = Tk()
        root.minsize(width=500, height=400)
        gui = Gui(root)
        root.mainloop()
