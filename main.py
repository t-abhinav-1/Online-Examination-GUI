from kivy.clock import Clock
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.screenmanager import ScreenManager, NoTransition, Screen
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.core.window import Window
import pandas as pd
import os


class TitleLabel(Label):
    def __init__(self, **kwargs):
        super(TitleLabel, self).__init__(**kwargs)
        self.text = "Online Examination GUI"
        self.font_size = '32sp'
        self.size_hint_y = None
        self.height = 70
        self.bold = True
        self.color = (0.8, 0.8, 0.8, 1)


class InputForm(GridLayout):
    def __init__(self, email_rollno_mapping, screen_manager, exam_button_callback, **kwargs):
        super(InputForm, self).__init__(**kwargs)
        self.cols = 1  # Single column layout
        self.screen_manager = screen_manager
        self.email_rollno_mapping = email_rollno_mapping
        self.exam_button_callback = exam_button_callback

        self.add_widget(TitleLabel())

        # Adjusted foreground_color to black for better visibility
        self.name_input = TextInput(hint_text="Enter your name", size_hint_y=None, height=40, multiline=False, foreground_color=(0, 0, 0, 1))
        self.email_input = TextInput(hint_text="Enter your email address", size_hint_y=None, height=40, multiline=False, foreground_color=(0, 0, 0, 1))
        self.rollno_input = TextInput(hint_text="Enter your roll number", size_hint_y=None, height=40, multiline=False, foreground_color=(0, 0, 0, 1))
        self.submit_button = Button(text="Submit", size_hint_y=None, height=40, background_color=(0.2, 0.6, 1, 1))
        self.submit_button.bind(on_press=self.verify_credentials)
        self.output_label = Label(text="", size_hint_y=None, height=40, color=(1, 0, 0, 1))

        self.add_widget(self.name_input)
        self.add_widget(self.email_input)
        self.add_widget(self.rollno_input)
        self.add_widget(self.submit_button)
        self.add_widget(self.output_label)

    def verify_credentials(self, instance):
        name = self.name_input.text
        email = self.email_input.text
        rollno = self.rollno_input.text

        print("Entered Email:", email)
        print("Entered Roll No:", rollno)

        if name and email and rollno:
            if email in self.email_rollno_mapping and self.email_rollno_mapping[email] == rollno:
                user_info = f"Name: {name}\nEmail: {email}\nRoll No: {rollno}"
                self.output_label.text = user_info

                # Set the user email for updating marks later
                exam_screen = self.screen_manager.get_screen("exam_screen")
                exam_screen.user_email = email

                self.exam_button_callback()

                self.screen_manager.current = "exam_screen"
            else:
                self.output_label.text = "Email and roll number do not match."
        else:
            self.output_label.text = "Please fill in all fields"


class ExamScreen(Screen):
    def __init__(self, question_file, key_file, excel_file, **kwargs):
        super(ExamScreen, self).__init__(**kwargs)
        self.question_file = question_file
        self.key_file = key_file
        self.excel_file = excel_file
        self.questions = []
        self.correct_answers = []
        self.selected_options = []
        self.user_email = None  # Store the user's email for updating marks

        self.current_question_index = 0
        self.marks = 0

        self.load_questions_and_answers()

        self.background_color = (0.5, 0.2, 0.9, 1)  # Greenish background color

    def load_questions_and_answers(self):
        try:
            with open(self.question_file, 'r') as q_file, open(self.key_file, 'r') as key_file:
                lines = q_file.readlines()
                self.correct_answers = key_file.read().splitlines()

                current_question = None
                current_options = []

                for line in lines:
                    line = line.strip()
                    if line:
                        if not current_question:
                            current_question = line
                        else:
                            current_options.append(line)

                        if len(current_options) == 4:
                            self.questions.append((current_question, current_options))
                            current_question = None
                            current_options = []

            self.correct_answers = [answer if answer else '-1' for answer in self.correct_answers]

            self.display_question()

        except FileNotFoundError:
            self.add_widget(Label(text="Error: Files not found"))

    def display_question(self):
        self.clear_widgets()

        main_layout = BoxLayout(orientation='vertical')
        main_layout.add_widget(TitleLabel())

        if 0 <= self.current_question_index < len(self.questions):
            question, options = self.questions[self.current_question_index]

            scroll_layout = GridLayout(cols=1, spacing=10, size_hint_y=None)
            scroll_layout.bind(minimum_height=scroll_layout.setter('height'))

            question_label = Label(text=f"Question {self.current_question_index + 1}: {question}",
                                   font_size='24sp', halign='left', valign='middle', size_hint_y=None, height=120,
                                   color=(1, 1, 1, 1))  # White color for text
            question_label.bind(size=question_label.setter('text_size'))
            scroll_layout.add_widget(question_label)

            options_layout = GridLayout(cols=1, spacing=10, size_hint_y=None)
            options_layout.bind(minimum_height=options_layout.setter('height'))

            for j, option in enumerate(options, start=1):
                option_button = Button(text=f"{option}", size_hint_y=None, height=36,
                                       background_color=(0.5, 0.7, 1, 1),  # Light blue background
                                       font_size='14sp')
                option_button.bind(on_press=self.select_option(j, option_button))
                options_layout.add_widget(option_button)

            scroll_layout.add_widget(options_layout)

            scroll_view = ScrollView(size_hint=(1, 1))
            scroll_view.add_widget(scroll_layout)
            main_layout.add_widget(scroll_view)

            if self.current_question_index == len(self.questions) - 1:
                submit_button = Button(text="Submit Exam", size_hint=(None, None), size=(200, 44),
                                       pos_hint={'center_x': 0.5},
                                       background_color=(0.8, 0.2, 0.2, 1))  # Red submit button
                submit_button.bind(on_press=self.submit_exam)
                main_layout.add_widget(submit_button)
            else:
                next_button = Button(text="Next", size_hint=(None, None), size=(100, 44), pos_hint={'center_x': 0.5},
                                     background_color=(0.2, 0.6, 0.2, 1))  # Green next button
                next_button.bind(on_press=self.next_question)
                main_layout.add_widget(next_button)

            self.add_widget(main_layout)

    def select_option(self, option_number, option_button):
        def on_select(instance):
            print(f"Selected Option {option_number}")
            if len(self.selected_options) > self.current_question_index:
                self.selected_options[self.current_question_index] = option_number
            else:
                self.selected_options.append(option_number)

            for button in instance.parent.children:
                button.background_color = (1, 1, 1, 1)
            instance.background_color = (0.5, 0.5, 1, 1)

        return on_select

    def next_question(self, instance):
        if 0 <= self.current_question_index < len(self.questions) - 1:
            self.current_question_index += 1
            self.display_question()

    def submit_exam(self, instance):
        if not self.questions or not self.correct_answers:
            self.clear_widgets()
            error_label = Label(text="Error: No questions or answers available", font_size='20sp')
            self.add_widget(error_label)
            return

        self.marks = sum(
            1 for i, selected_option in enumerate(self.selected_options) if 0 <= i < len(self.correct_answers) and
            selected_option == int(self.correct_answers[i]))

        # Update marks in Excel
        self.update_marks_in_excel()

        self.clear_widgets()
        result_label = Label(text=f"Your marks: {self.marks}", font_size='20sp')
        self.add_widget(result_label)

        Clock.schedule_once(self.close_app, 4)

    def update_marks_in_excel(self):
        try:
            df = pd.read_excel(self.excel_file)

            # Update 'Marks' column based on user's email
            if self.user_email is not None:
                df.loc[df['Email'] == self.user_email, 'Marks'] = self.marks

                # Save the updated DataFrame to Excel
                df.to_excel(self.excel_file, index=False)

        except FileNotFoundError:
            print("Error storing marks in Excel: Excel file not found.")
        except Exception as e:
            print("Error storing marks in Excel:", e)

    def close_app(self, dt):
        App.get_running_app().stop()


class UserInfoApp(App):
    def __init__(self, excel_file, question_file, key_file, **kwargs):
        super(UserInfoApp, self).__init__(**kwargs)
        self.excel_file = excel_file
        self.question_file = question_file
        self.key_file = key_file
        self.exam_button_enabled = False

    def build(self):
        screen_manager = ScreenManager(transition=NoTransition())

        Window.fullscreen = 'auto'

        login_screen = Screen(name="login_screen")
        email_rollno_mapping = self.load_data_from_excel()

        def enable_exam_button():
            self.exam_button_enabled = True

        input_form = InputForm(email_rollno_mapping, screen_manager, enable_exam_button)
        login_screen.add_widget(input_form)
        screen_manager.add_widget(login_screen)

        exam_screen = ExamScreen(name="exam_screen", question_file=self.question_file, key_file=self.key_file,
                                 excel_file=self.excel_file)
        screen_manager.add_widget(exam_screen)

        return screen_manager

    def load_data_from_excel(self):
        email_rollno_mapping = {}

        try:
            df = pd.read_excel(self.excel_file)

            # Print column names for debugging
            print("Column Names in Excel File:", df.columns.tolist())

            # Assuming 'Email' and 'Roll No' are column names in your Excel file
            email_column = 'Email'
            rollno_column = 'Roll No'

            if email_column in df.columns and rollno_column in df.columns:
                email_rollno_mapping = {row[email_column]: row[rollno_column] for _, row in df.iterrows()}
                print("Loaded Email-RollNo Mapping:", email_rollno_mapping)
            else:
                print(f"Error: '{email_column}' or '{rollno_column}' columns not found in Excel file.")

        except FileNotFoundError:
            print("Error: Excel file not found.")
        except Exception as e:
            print("Error loading data from Excel:", e)

        return email_rollno_mapping


if __name__ == '__main__':
    # Assuming 'assets' directory is in the same directory as your script

    print("Current directory:", os.getcwd())  # Print current working directory
    print("Excel file path:", os.path.abspath("Verify_Credentials.xlsx"))  # Print absolute path to Excel file

    assets_dir = os.path.join(os.path.dirname(__file__), "assets")
    excel_file = os.path.join(assets_dir, "Verify_Credentials.xlsx")

    # excel_file = r"C:\Users\HP\PycharmProjects\Online-Examination-GUI\assets\Verify_Credentilas.xlsx"
    question_file = "que_and_options.txt"
    key_file = "key.txt"
    UserInfoApp(excel_file, question_file, key_file).run()
