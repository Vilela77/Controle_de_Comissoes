import kivy
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.spinner import Spinner
import pandas as pd
import os
import subprocess
from kivy.uix.popup import Popup
from kivy.uix.label import Label


class CommissionCalculator(App):
    def build(self):
        self.layout = BoxLayout(orientation='vertical', padding=3, spacing=3)
        self.client_code_input = TextInput(hint_text='Código do Cliente', font_size=40, hint_text_color=(10, 10, 10, 10),
                                           foreground_color=(10, 10, 10, 10), background_color=(0, 0, 0.2, 1))
        self.client_name_input = TextInput(hint_text='Nome', font_size=40, hint_text_color=(10, 10, 10, 10),
                                           foreground_color=(10, 10, 10, 10), background_color=(0, 0, 0.2, 1))
        self.service_type_spinner = Spinner(text='Tipo de Serviço', values=('Visita Tecnica', 'Recolhimento', 'Venda Nova'),
                                            font_size=28, background_color=(0, 0, 1, 1))
        self.salvar_button = Button(text='Salvar', font_size=28, background_color=(0, 0, 1, 1))
        self.abrir_planilha_button = Button(text='Abrir Planilha', font_size=28, background_color=(0, 0, 1, 1))
        self.salvar_button.bind(on_press=self.salvar_dados)
        self.abrir_planilha_button.bind(on_press=self.abrir_planilha)

        self.layout.add_widget(self.client_code_input)
        self.layout.add_widget(self.client_name_input)
        self.layout.add_widget(self.service_type_spinner)
        self.layout.add_widget(self.salvar_button)
        self.layout.add_widget(self.abrir_planilha_button)

        return self.layout

    def salvar_dados(self, instance):
        client_code = self.client_code_input.text
        client_name = self.client_name_input.text
        service_type = self.service_type_spinner.text

        # Verifica se o usuário digitou um código e um nome
        if not client_code or not client_name:
            popup = Popup(title='Aviso', content=Label(text='Por favor, preencha o código do cliente e o nome.'),
                        size_hint=(None, None), size=(400, 200))
            popup.open()
            return

        # Verifica se o usuário selecionou uma opção válida no Spinner
        if service_type == 'Tipo de Serviço':
            popup = Popup(title='Aviso', content=Label(text='Por favor, selecione um tipo de serviço válido.'),
                      size_hint=(None, None), size=(400, 200))
            popup.open()
            return

        # Lógica para calcular a comissão com base no tipo de serviço
        commission = 0
        if service_type == 'Visita Tecnica':
            commission = 8
        elif service_type == 'Recolhimento':
            commission = 10
        elif service_type == 'Venda Nova':
            commission = 10

        # Verifica se o arquivo Excel já existe
        file_exists = os.path.isfile('commissions.xlsx')

        if file_exists:
        # Carregar o DataFrame existente
            df = pd.read_excel('commissions.xlsx')
        # Calcular o novo total
            total = df['Comissão'].sum() + commission
        else:
        # Se o arquivo não existir, cria um novo DataFrame e define o total inicial
            df = pd.DataFrame(
            columns=['Código do Cliente', 'Primeiro Nome', 'Tipo de Serviço', 'Comissão', 'Total'])
            total = commission

        # Adicionar as novas informações ao DataFrame
        new_data = {'Código do Cliente': client_code, 'Primeiro Nome': client_name, 'Tipo de Serviço': service_type,
                'Comissão': commission, 'Total': total}
        df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)

        # Salvar as informações em um arquivo Excel
        df.to_excel('commissions.xlsx', index=False)
        print("Informações salvas com sucesso!")

        # Limpar os campos de entrada
        self.client_code_input.text = ''
        self.client_name_input.text = ''
        self.service_type_spinner.text = 'Tipo de Serviço'


    def abrir_planilha(self, instance):
        # Verifica se o arquivo Excel existe
        if os.path.isfile('commissions.xlsx'):
            # Abre o arquivo Excel usando o programa padrão do sistema operacional
            if kivy.utils.platform == 'win':
                os.startfile('commissions.xlsx')
            elif kivy.utils.platform == 'macosx':
                subprocess.call(['open', 'commissions.xlsx'])
            else:
                subprocess.call(['xdg-open', 'commissions.xlsx'])
        else:
            print("O arquivo commissions.xlsx não existe.")


if __name__ == '__main__':
    CommissionCalculator().run()
