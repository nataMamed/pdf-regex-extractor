from PySimpleGUI import PySimpleGUI as sg
import pandas as pd
from PyPDF2 import PdfReader
import re
import os

class GUI:

    sg.theme('Reddit')

    welcome = """Bem-vindo a ferramenta de extração de texto de PDF's"""
    source = ''
    destiny = ''
    max_progress_bar = 1000
    layout = [
        [sg.Text(welcome, enable_events=True)],
        [sg.Text('Selecione o arquivo'), sg.Button('selectFile'), sg.Text('Selecione onde o arquivo será guardado'), sg.Button('selectFolder')],
        [],

        [sg.Text(' '*50, key='warning', text_color='red')],
        [sg.Text('Arquivo a ser procurado:'), sg.Text(source, key='file')],
        [sg.Text('Onde Será guardado:'), sg.Text(source, key='destiny')],

        [sg.Text('Texto a ser procurado'), sg.Input(key='pattern')],

        [sg.Button('Procurar'), ],
        [sg.Text('Numero de páginas:'), sg.Text('', key='pages')],
        [sg.ProgressBar(max_progress_bar, orientation='h', key='progress_bar')],

        [sg.Text('Status:'), sg.Text(
            'Parado', auto_size_text=True, key='status')],
 
    ]

    window = sg.Window('PDF Extractor', layout)


    def extract_data_from_pdf(self, path, pattern):

        related_words = {"page":[], "word":[]}
        reader = PdfReader(path)
        n_pages = len(reader.pages)
        

        self.window['pages'].update(f'{n_pages}')
        progess_counter = 0
        counter_to_add = self.max_progress_bar / n_pages

        for page_counter, page in enumerate(reader.pages):
            self.window['progress_bar'].update(current_count=int(progess_counter))
            
            
            try:
                text = page.extract_text().lower()
                match_texts = re.findall(f".*{pattern}.*", text)
                # print(match_texts)
                if match_texts:
                    for item in match_texts:
                        related_words["page"].append(page_counter+1)
                        related_words["word"].append(item.strip())
                page_counter += 1
                progess_counter += counter_to_add
                
            except Exception as err:
                pass
                # print("Error")

        df = pd.DataFrame(related_words)

        return df

    def run(self):

        while True:
            events, values = self.window.read()
            if events == sg.WIN_CLOSED:
                break
            

            if events == 'selectFile':

                self.source = sg.PopupGetFile('Favor selecionar um arquivo PDF',
                default_path='', default_extension='.xlsx',
                save_as=False, file_types=(('PDF', '.pdf'),), 
                no_window=False, font=None, no_titlebar=False,
                grab_anywhere=False)

                self.window['file'].update(self.source)
                if self.source != '':
                    self.window['warning'].update('')
            
            if events == 'selectFolder':
                self.destiny = sg.PopupGetFolder('Favor selecionar uma pasta',
                default_path='', 
                no_window=False, font=None, no_titlebar=False,
                grab_anywhere=False)
                self.window['destiny'].update(self.destiny)
                if self.destiny != '':
                    self.window['warning'].update('')

               
            if events == 'Procurar':
                if self.source != '' and values['pattern'] and self.destiny:
                    try:
                        if values['pattern']:
                            self.window['status'].update('Rodando..')
                            result = self.extract_data_from_pdf(path=self.source, pattern=values['pattern'].strip().lower())
                            self.window['status'].update('Terminado..')
                            save_as = self.destiny + '/' +  self.source.split('/')[-1].replace('pdf', 'csv')
                            result.to_csv(save_as, index=False)
 
                    except Exception as err:
                        # print(err)
                        self.window['warning'].update('Endereço não encontrado')

                elif not self.source:
                    self.window['warning'].update('Selecione o arquivo a ser utilizado')  

                elif not self.destiny:
                    self.window['warning'].update('Selecione o destino do arquivo')    

                elif not values['pattern']:
                    self.window['warning'].update('Digite o texto a ser procurado')

                else:
                    self.window['warning'].update('')

        self.window.close()



if __name__ == '__main__':
    gui = GUI()
    gui.run()