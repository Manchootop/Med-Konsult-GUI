from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.filechooser import FileChooserIconView
from kivy.uix.popup import Popup
from docx import Document
import os


class ScreeningFormApp(App):

    def build(self):
        # Layout for the form
        layout = BoxLayout(orientation='vertical', padding=10, spacing=10)

        # Creating input fields
        self.names_input = TextInput(hint_text="Enter Patient's Full Name", multiline=False)
        self.dob_input = TextInput(hint_text="Enter Date of Birth (DD.MM.YYYY)", multiline=False)
        self.doctor_input = TextInput(hint_text="Enter Doctor's Name", multiline=False)
        self.coordinator_input = TextInput(hint_text="Enter Coordinator's Name", multiline=False)
        self.screening_time_input = TextInput(hint_text="Enter Screening Time (HH:MM)", multiline=False)
        self.age_input = TextInput(hint_text="Enter Patient's Age", multiline=False)
        self.id_input = TextInput(hint_text="Enter Patient's ID Number", multiline=False)

        # Submit button
        submit_button = Button(text="Generate Document", on_press=self.generate_document)

        # Button to open file chooser dialog for Word file
        preview_button = Button(text="Preview Word Document", on_press=self.open_file_chooser)

        # Label to display the generated document
        self.result_label = Label(size_hint_y=None, height=200, text='Generated Document will appear here...')

        # Scroll view for the result (in case it's long)
        scroll_view = ScrollView(size_hint=(1, None), size=(400, 200))
        scroll_view.add_widget(self.result_label)

        # Adding widgets to the layout
        layout.add_widget(self.names_input)
        layout.add_widget(self.dob_input)
        layout.add_widget(self.doctor_input)
        layout.add_widget(self.coordinator_input)
        layout.add_widget(self.screening_time_input)
        layout.add_widget(self.age_input)
        layout.add_widget(self.id_input)
        layout.add_widget(submit_button)
        layout.add_widget(preview_button)
        layout.add_widget(scroll_view)

        return layout

    def generate_document(self, instance):
        # Retrieve the input values
        names = self.names_input.text
        dob = self.dob_input.text
        doctor = self.doctor_input.text
        coordinator = self.coordinator_input.text
        screening_time = self.screening_time_input.text
        age = self.age_input.text
        patient_id = self.id_input.text

        # Document template with placeholders
        document_template = f"""
        СПОНСОР: ModernaTX, Inc.
        ЦЕНТЪР: BG004                                                                                              ДАТА: 12.11.2024
        ПРОТОКОЛ №: mRNA-1010-P304                               Пациент: {names}
        ГЛАВЕН ИЗСЛЕДОВАТЕЛ: Проф. Д-р М. Цекова            Номер пациент: {patient_id}
        ......
        Пациентът бе информиран, че ако има партньорка в детеродна възраст трябва да използват подходящи методи за контрол на раждаемостта.
        Пациентът отговаря на всички включващи и няма нито един от изключващите критерии към момента.
        Пациентът бе скриниран в {screening_time} ч. и му бе назначен номер: {patient_id}.
        Нежелани събития по време на визитата не се регистрират.
        Визитата бе изготвена под диктовката на {doctor} от координатор по клиничното проучване {coordinator}.
        """

        # Display the generated document
        self.result_label.text = document_template

        # Save the generated document to a Word file
        self.save_document(patient_id, document_template)

    def save_document(self, patient_id, document_text):
        # Create a new Document object
        doc = Document()

        # Add the text content to the document
        doc.add_paragraph(document_text)

        # Define the file path to save the document
        file_name = f"скрининг_{patient_id}.docx"

        # Ensure the file is saved in the current working directory
        file_path = os.path.join(os.getcwd(), file_name)

        # Save the document
        doc.save(file_path)
        print(f"Document saved as {file_path}")

    def open_file_chooser(self, instance):
        # Create file chooser dialog for the user to select a Word document
        file_chooser = FileChooserIconView()
        file_chooser.filters = ['*.docx']  # Show only .docx files

        # Create a popup to show the file chooser
        popup = Popup(title="Select Word Document", content=file_chooser, size_hint=(0.8, 0.8))

        # Add an "Open" button to the file chooser
        open_button = Button(text="Open", on_press=lambda x: self.load_word_file(file_chooser.selection, popup))
        popup.content.add_widget(open_button)

        popup.open()

    def load_word_file(self, selected, popup):
        if selected:
            word_file = selected[0]
            # Extract text from the selected Word file
            document_text = self.extract_text_from_word(word_file)
            self.result_label.text = document_text  # Display extracted text
            popup.dismiss()  # Close the popup after selecting the file

    def extract_text_from_word(self, word_file):
        # Extract text from the Word file using python-docx
        document = Document(word_file)
        full_text = []
        for para in document.paragraphs:
            full_text.append(para.text)
        return '\n'.join(full_text)


if __name__ == "__main__":
    ScreeningFormApp().run()
