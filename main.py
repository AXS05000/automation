from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.boxlayout import BoxLayout
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from kivy.uix.popup import Popup
from kivy.uix.filechooser import FileChooserIconView

class MyApp(App):

    def build(self):
        layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        
        open_browser_button = Button(text="Abrir Navegador", on_press=self.open_browser)
        navigate_button = Button(text="Acessar Página", on_press=self.navigate_to_page)
        execute_script_button = Button(text="Executar Script", on_press=self.execute_script)

        layout.add_widget(open_browser_button)
        layout.add_widget(navigate_button)
        layout.add_widget(execute_script_button)
        return layout

    def configure_firefox(self):
        firefox_options = webdriver.FirefoxOptions()
        firefox_options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        return firefox_options

    def open_browser(self, instance):
        firefox_options = self.configure_firefox()
        self.browser = webdriver.Firefox(options=firefox_options)
        self.browser.get('https://app.folhacerta.com/admin/login')  # URL da tela de login do sistema

    def navigate_to_page(self, instance):
        self.browser.get('https://app.folhacerta.com/admin/JornadaRegraAlocacao?jornadaRegraAlocacao.Id=0')  # URL da página específica
        element_present = EC.presence_of_element_located((By.ID, 'content'))
        WebDriverWait(self.browser, 10).until(element_present)

    def execute_script(self, instance):
        script = """
        // Define os valores
        var jornadaRegraIdValue = "4300";  // ID da escala
        var dataInicioValue = "04/10/2023"; // Data de vigência inicial
        var usuarioIdValue = "99176"; // ID do ALESSANDRE HEMERSON DOS REIS

        // Preenche os campos do formulário
        document.querySelector('[name="jornadaRegra_id"]').value = jornadaRegraIdValue;
        document.querySelector('[name="jornadaRegraAlocacao.DataInicio"]').value = dataInicioValue;

        // Seleciona o usuário
        var usuarioSelect = document.querySelector('#select-multi-usuarios');
        for (var i = 0; i < usuarioSelect.options.length; i++) {
            if (usuarioSelect.options[i].value === usuarioIdValue) {
                usuarioSelect.options[i].selected = true;
                break;
            }
        }

        // Envia o formulário
        document.querySelector('form.form-horizontal').submit();
        """

        self.browser.execute_script(script)

if __name__ == '__main__':
    MyApp().run()
