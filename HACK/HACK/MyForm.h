#pragma once
#include <zip.h>
#include <string>
#include <vector>
#include <stdexcept>
#include "Doczip.h"

bool processDocx(const std::string& inputPath, const std::string& outputPath) {
    std::string tempDir = "unzipped_temp";
    std::string docXml = tempDir + "/word/document.xml";

    if (!extractDocx(inputPath, tempDir)) {
        std::cerr << "Распаковка не удалась!" << std::endl;
        return false;
    }

    if (!fs::exists(docXml)) {
        std::cerr << "Файл document.xml не найден!" << std::endl;
        return false;
    }

    return true;
}

namespace HACK {

    using namespace System;
    using namespace System::Windows::Forms;
    using namespace System::Drawing;
    using namespace msclr::interop;


    public ref class MyForm : public System::Windows::Forms::Form {
    public:
        MyForm(void) {
            InitializeComponent();
            isDarkTheme = false;
            ApplyTheme();
        }

    protected:
        ~MyForm() {
            if (components)
                delete components;
        }

    public:
        bool isDarkTheme;
    private:

        System::Windows::Forms::Button^ button1;
        System::Windows::Forms::Label^ label1;
        System::Windows::Forms::TextBox^ textBox1;
        System::Windows::Forms::Panel^ topPanel;
        System::Windows::Forms::Label^ buttonTheme;
    private: System::Windows::Forms::TextBox^ outputNameBox;
    private: System::Windows::Forms::Label^ label2;
    private: System::Windows::Forms::TextBox^ textBox2;
    private: System::Windows::Forms::TextBox^ textBox3;
    private: System::Windows::Forms::TextBox^ textBox4;
    private: System::Windows::Forms::TextBox^ textBox5;
    private: System::Windows::Forms::CheckBox^ checkBox1;


           System::ComponentModel::Container^ components;

           void ApplyTheme() {
               if (isDarkTheme) {
                   this->BackColor = Color::FromArgb(30, 30, 30);
                   this->ForeColor = Color::White;
                   textBox1->BackColor = Color::FromArgb(45, 45, 45);
                   textBox1->ForeColor = Color::White;
                   outputNameBox->BackColor = Color::FromArgb(45, 45, 45);
                   outputNameBox->ForeColor = Color::White;
                   button1->BackColor = Color::FromArgb(45, 45, 45);
                   button1->ForeColor = Color::White;
                   buttonTheme->BackColor = Color::Black;
                   buttonTheme->ForeColor = Color::White;
                   topPanel->BackColor = Color::Black;
               }
               else {
                   this->BackColor = Color::WhiteSmoke;
                   this->ForeColor = Color::Black;
                   textBox1->BackColor = Color::White;
                   textBox1->ForeColor = Color::Black;
                   outputNameBox->BackColor = Color::White;
                   outputNameBox->ForeColor = Color::Black;
                   button1->BackColor = Color::RoyalBlue;
                   button1->ForeColor = Color::White;
                   buttonTheme->BackColor = Color::RoyalBlue;
                   buttonTheme->ForeColor = Color::Black;
                   topPanel->BackColor = Color::RoyalBlue;
               }
           }

#pragma region Windows Form Designer generated code
           void InitializeComponent(void) {
               this->button1 = (gcnew System::Windows::Forms::Button());
               this->label1 = (gcnew System::Windows::Forms::Label());
               this->textBox1 = (gcnew System::Windows::Forms::TextBox());
               this->topPanel = (gcnew System::Windows::Forms::Panel());
               this->buttonTheme = (gcnew System::Windows::Forms::Label());
               this->outputNameBox = (gcnew System::Windows::Forms::TextBox());
               this->label2 = (gcnew System::Windows::Forms::Label());
               this->textBox2 = (gcnew System::Windows::Forms::TextBox());
               this->textBox3 = (gcnew System::Windows::Forms::TextBox());
               this->textBox4 = (gcnew System::Windows::Forms::TextBox());
               this->textBox5 = (gcnew System::Windows::Forms::TextBox());
               this->checkBox1 = (gcnew System::Windows::Forms::CheckBox());
               this->topPanel->SuspendLayout();
               this->SuspendLayout();
               // 
               // button1
               // 
               this->button1->FlatStyle = System::Windows::Forms::FlatStyle::Flat;
               this->button1->Location = System::Drawing::Point(337, 348);
               this->button1->Name = L"button1";
               this->button1->Size = System::Drawing::Size(140, 35);
               this->button1->TabIndex = 4;
               this->button1->Text = L"Форматировать";
               this->button1->Click += gcnew System::EventHandler(this, &MyForm::button1_Click);
               // 
               // label1
               // 
               this->label1->AutoSize = true;
               this->label1->Font = (gcnew System::Drawing::Font(L"Segoe UI", 10));
               this->label1->Location = System::Drawing::Point(333, 109);
               this->label1->Name = L"label1";
               this->label1->Size = System::Drawing::Size(144, 19);
               this->label1->TabIndex = 3;
               this->label1->Text = L"Имя входного файла:";
               // 
               // textBox1
               // 
               this->textBox1->Font = (gcnew System::Drawing::Font(L"Segoe UI", 10));
               this->textBox1->Location = System::Drawing::Point(356, 131);
               this->textBox1->Multiline = true;
               this->textBox1->Name = L"textBox1";
               this->textBox1->Size = System::Drawing::Size(101, 22);
               this->textBox1->TabIndex = 2;
               // 
               // topPanel
               // 
               this->topPanel->Controls->Add(this->buttonTheme);
               this->topPanel->Dock = System::Windows::Forms::DockStyle::Top;
               this->topPanel->Location = System::Drawing::Point(0, 0);
               this->topPanel->Name = L"topPanel";
               this->topPanel->Size = System::Drawing::Size(854, 60);
               this->topPanel->TabIndex = 1;
               // 
               // buttonTheme
               // 
               this->buttonTheme->BackColor = System::Drawing::Color::White;
               this->buttonTheme->Location = System::Drawing::Point(822, 9);
               this->buttonTheme->Name = L"buttonTheme";
               this->buttonTheme->Size = System::Drawing::Size(20, 20);
               this->buttonTheme->TabIndex = 0;
               this->buttonTheme->Text = L"☼";
               this->buttonTheme->Click += gcnew System::EventHandler(this, &MyForm::buttonTheme_Click);
               // 
               // outputNameBox
               // 
               this->outputNameBox->Location = System::Drawing::Point(356, 183);
               this->outputNameBox->Name = L"outputNameBox";
               this->outputNameBox->Size = System::Drawing::Size(100, 23);
               this->outputNameBox->TabIndex = 5;
               // 
               // label2
               // 
               this->label2->AutoSize = true;
               this->label2->Font = (gcnew System::Drawing::Font(L"Segoe UI", 10));
               this->label2->Location = System::Drawing::Point(333, 162);
               this->label2->Name = L"label2";
               this->label2->Size = System::Drawing::Size(154, 19);
               this->label2->TabIndex = 6;
               this->label2->Text = L"Имя выходного файла:";
               // 
               // textBox2
               // 
               this->textBox2->Font = (gcnew System::Drawing::Font(L"Segoe UI", 10));
               this->textBox2->Location = System::Drawing::Point(691, 131);
               this->textBox2->Multiline = true;
               this->textBox2->Name = L"textBox2";
               this->textBox2->Size = System::Drawing::Size(101, 22);
               this->textBox2->TabIndex = 7;
               // 
               // textBox3
               // 
               this->textBox3->Font = (gcnew System::Drawing::Font(L"Segoe UI", 10));
               this->textBox3->Location = System::Drawing::Point(691, 181);
               this->textBox3->Multiline = true;
               this->textBox3->Name = L"textBox3";
               this->textBox3->Size = System::Drawing::Size(101, 22);
               this->textBox3->TabIndex = 8;
               // 
               // textBox4
               // 
               this->textBox4->Font = (gcnew System::Drawing::Font(L"Segoe UI", 10));
               this->textBox4->Location = System::Drawing::Point(691, 224);
               this->textBox4->Multiline = true;
               this->textBox4->Name = L"textBox4";
               this->textBox4->Size = System::Drawing::Size(101, 22);
               this->textBox4->TabIndex = 9;
               // 
               // textBox5
               // 
               this->textBox5->Font = (gcnew System::Drawing::Font(L"Segoe UI", 10));
               this->textBox5->Location = System::Drawing::Point(691, 264);
               this->textBox5->Multiline = true;
               this->textBox5->Name = L"textBox5";
               this->textBox5->Size = System::Drawing::Size(101, 22);
               this->textBox5->TabIndex = 10;
               // 
               // checkBox1
               // 
               this->checkBox1->AutoSize = true;
               this->checkBox1->Location = System::Drawing::Point(356, 214);
               this->checkBox1->Name = L"checkBox1";
               this->checkBox1->Size = System::Drawing::Size(82, 19);
               this->checkBox1->TabIndex = 11;
               this->checkBox1->Text = L"checkBox1";
               this->checkBox1->UseVisualStyleBackColor = true;
               // 
               // MyForm
               // 
               this->ClientSize = System::Drawing::Size(854, 416);
               this->Controls->Add(this->checkBox1);
               this->Controls->Add(this->textBox5);
               this->Controls->Add(this->textBox4);
               this->Controls->Add(this->textBox3);
               this->Controls->Add(this->textBox2);
               this->Controls->Add(this->label2);
               this->Controls->Add(this->outputNameBox);
               this->Controls->Add(this->topPanel);
               this->Controls->Add(this->textBox1);
               this->Controls->Add(this->label1);
               this->Controls->Add(this->button1);
               this->Font = (gcnew System::Drawing::Font(L"Segoe UI", 9));
               this->FormBorderStyle = System::Windows::Forms::FormBorderStyle::FixedDialog;
               this->MaximizeBox = false;
               this->Name = L"MyForm";
               this->StartPosition = System::Windows::Forms::FormStartPosition::CenterScreen;
               this->Text = L"DOCX Reader";
               this->topPanel->ResumeLayout(false);
               this->ResumeLayout(false);
               this->PerformLayout();

           }
#pragma endregion

    private: System::Void buttonTheme_Click(System::Object^ sender, System::EventArgs^ e) {
        isDarkTheme = !isDarkTheme;
        ApplyTheme();
    }
    private: System::Void button1_Click(System::Object^ sender, System::EventArgs^ e) {
        String^ inputPathStr = textBox1->Text;
        String^ outputNameStr = outputNameBox->Text;

        if (String::IsNullOrWhiteSpace(inputPathStr)) {
            MessageBox::Show("Введите путь к .docx файлу", "Ошибка");
            return;
        }

        if (String::IsNullOrWhiteSpace(outputNameStr)) {
            MessageBox::Show("Введите имя выходного файла", "Ошибка");
            return;
        }

        std::string inputPath = msclr::interop::marshal_as<std::string>(inputPathStr);
        std::string outputPath = msclr::interop::marshal_as<std::string>(outputNameStr) + ".docx";
        std::string tempDir = "unzipped_temp";
        std::string docXml = tempDir + "/word/document.xml";

        try {
            if (!extractDocx(inputPath, tempDir)) {
                MessageBox::Show("Ошибка при распаковке файла", "Ошибка");
                return;
            }

            if (!System::IO::File::Exists(gcnew String(docXml.c_str()))) {
                MessageBox::Show("Файл document.xml не найден", "Ошибка");
                return;
            }
            String^ str1 = textBox2->Text;
            String^ str2 = textBox3->Text;
            String^ str3 = textBox4->Text;
            String^ str4 = textBox5->Text;

            std::string s1 = to_utf8(str1);
            std::string s2 = to_utf8(str2);
            std::string s3 = to_utf8("/" + str3);
            std::string s4 = to_utf8("Студент группы " + str4);

            formatDocumentXml(docXml);
            formatHeadingsInFile(docXml);
            addTitlePage(docXml, s1, s2, s3, s4);
            MessageBox::Show("Файл успешно отформатирован и сохранён как:\n" + outputNameStr + ".docx", "Готово");
        }
        catch (const std::exception& ex) {
            MessageBox::Show(gcnew String(ex.what()), "Ошибка");
        }
    }

    };
}