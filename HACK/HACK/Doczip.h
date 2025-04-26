#include <iostream>
#include <filesystem>
#include <fstream>
#include <zip.h>
#include <tinyxml2.h>
#include<string>
#include <msclr/marshal.h>
#include <msclr/marshal_cppstd.h>
namespace fs = std::filesystem;
using namespace tinyxml2;

std::string to_utf8(System::String^ str) {
    System::Text::Encoding^ utf8 = System::Text::Encoding::UTF8;
    array<System::Byte>^ bytes = utf8->GetBytes(str);
    pin_ptr<System::Byte> pinned = &bytes[0];
    return std::string(reinterpret_cast<char*>(pinned), bytes->Length);
}


bool extractDocx(const std::string& inputPath, const std::string& outDir) {
    int err = 0;
    zip* z = zip_open(inputPath.c_str(), ZIP_RDONLY, &err);
    if (!z) {
        std::cerr << "Не удалось открыть docx" << std::endl;
        return false;
    }
    zip_int64_t num = zip_get_num_entries(z, 0);
    for (zip_uint64_t i = 0; i < num; ++i) {
        const char* name = zip_get_name(z, i, 0);
        if (!name) continue;
        zip_file* f = zip_fopen_index(z, i, 0);
        if (!f) continue;
        std::string outPath = outDir + "/" + name;
        fs::create_directories(fs::path(outPath).parent_path());
        std::ofstream out(outPath, std::ios::binary);
        char buf[4096];
        zip_int64_t n;
        while ((n = zip_fread(f, buf, sizeof(buf))) > 0)
            out.write(buf, n);
        zip_fclose(f);
        out.close();
    }
    zip_close(z);
    return true;
}

void formatDocumentXml(const std::string& xmlPath) {
    tinyxml2::XMLDocument doc;
    if (doc.LoadFile(xmlPath.c_str()) != XML_SUCCESS) {
        std::cerr << "Ошибка загрузки XML" << std::endl;
        return;
    }
    XMLElement* root = doc.FirstChildElement("w:document");
    if (!root) return;
    XMLElement* body = root->FirstChildElement("w:body");
    for (XMLElement* para = body->FirstChildElement("w:p"); para; para = para->NextSiblingElement("w:p")) {
        for (XMLElement* run = para->FirstChildElement("w:r"); run; run = run->NextSiblingElement("w:r")) {
            XMLElement* rPr = run->FirstChildElement("w:rPr");
            if (!rPr) rPr = run->InsertFirstChild(doc.NewElement("w:rPr"))->ToElement();
            XMLElement* fonts = rPr->FirstChildElement("w:rFonts");
            if (!fonts) fonts = rPr->InsertNewChildElement("w:rFonts");
            fonts->SetAttribute("w:ascii", "Times New Roman");
            fonts->SetAttribute("w:hAnsi", "Times New Roman");
            XMLElement* size = rPr->FirstChildElement("w:sz");
            if (!size) size = rPr->InsertNewChildElement("w:sz");
            size->SetAttribute("w:val", "28");
            XMLElement* color = rPr->FirstChildElement("w:color");
            if (!color) color = rPr->InsertNewChildElement("w:color");
            color->SetAttribute("w:val", "000000");
        }
        XMLElement* pPr = para->FirstChildElement("w:pPr");
        if (!pPr) pPr = para->InsertFirstChild(doc.NewElement("w:pPr"))->ToElement();
        XMLElement* spacing = pPr->FirstChildElement("w:spacing");
        if (!spacing) spacing = pPr->InsertNewChildElement("w:spacing");
        spacing->SetAttribute("w:line", "360");
        spacing->SetAttribute("w:lineRule", "auto");
    }
    doc.SaveFile(xmlPath.c_str());
}

void formatHeadingsInFile(const std::string& filePath) {
    tinyxml2::XMLDocument doc;
    tinyxml2::XMLError eResult = doc.LoadFile(filePath.c_str());
    if (eResult != tinyxml2::XML_SUCCESS) {
        std::cerr << "Ошибка при загрузке файла: " << filePath << std::endl;
        return;
    }
    tinyxml2::XMLElement* body = doc.FirstChildElement("w:document")->FirstChildElement("w:body");
    for (tinyxml2::XMLElement* p = body->FirstChildElement("w:p"); p != nullptr; p = p->NextSiblingElement("w:p")) {
        tinyxml2::XMLElement* style = p->FirstChildElement("w:pPr")->FirstChildElement("w:pStyle");
        if (style && style->Attribute("w:val")) {
            std::string styleVal = style->Attribute("w:val");
            if (styleVal == "1" || styleVal == "2" || styleVal == "3") {
                tinyxml2::XMLElement* spacing = p->FirstChildElement("w:pPr")->FirstChildElement("w:spacing");
                if (!spacing) {
                    spacing = doc.NewElement("w:spacing");
                    p->FirstChildElement("w:pPr")->InsertEndChild(spacing);
                }
                spacing->SetAttribute("w:before", "480");
                spacing->SetAttribute("w:after", "480");
                tinyxml2::XMLElement* ind = p->FirstChildElement("w:pPr")->FirstChildElement("w:ind");
                if (!ind) {
                    ind = doc.NewElement("w:ind");
                    p->FirstChildElement("w:pPr")->InsertEndChild(ind);
                }
                ind->SetAttribute("w:left", "35");
            }
        }
    }
    eResult = doc.SaveFile(filePath.c_str());
    if (eResult != tinyxml2::XML_SUCCESS) {
        std::cerr << "Ошибка при сохранении файла: " << filePath << std::endl;
        return;
    }
}

void addTitlePage(const std::string& filePath, const std::string& topic,
    const std::string& docType, const std::string& studentName,
    const std::string& group) {
    using namespace tinyxml2;
    tinyxml2::XMLDocument doc;
    if (doc.LoadFile(filePath.c_str()) != XML_SUCCESS) {
        std::cerr << "Не удалось загрузить файл: " << filePath << std::endl;
        return;
    }

    XMLElement* document = doc.FirstChildElement("w:document");
    if (!document) {
        std::cerr << "Не найден <w:document>" << std::endl;
        return;
    }

    XMLElement* body = document->FirstChildElement("w:body");
    if (!body) {
        std::cerr << "Не найден <w:body>" << std::endl;
        return;
    }

    auto createParagraph = [&](const std::string& text, bool isLast = false) -> XMLElement* {
        XMLElement* p = doc.NewElement("w:p");

        XMLElement* pPr = doc.NewElement("w:pPr");
        XMLElement* jc = doc.NewElement("w:jc");
        jc->SetAttribute("w:val", "center");
        pPr->InsertEndChild(jc);
        p->InsertEndChild(pPr);

        XMLElement* r = doc.NewElement("w:r");
        XMLElement* rPr = doc.NewElement("w:rPr");

        XMLElement* rFonts = doc.NewElement("w:rFonts");
        rFonts->SetAttribute("w:ascii", "Times New Roman");
        rFonts->SetAttribute("w:hAnsi", "Times New Roman");
        rPr->InsertEndChild(rFonts);

        XMLElement* sz = doc.NewElement("w:sz");
        sz->SetAttribute("w:val", "28");
        rPr->InsertEndChild(sz);

        r->InsertEndChild(rPr);

        XMLElement* t = doc.NewElement("w:t");
        t->SetText(text.c_str());
        r->InsertEndChild(t);

        if (isLast) {
            XMLElement* br = doc.NewElement("w:br");
            br->SetAttribute("w:type", "page");
            r->InsertEndChild(br);

            XMLElement* sectPr = doc.NewElement("w:sectPr");

            XMLElement* titlePg = doc.NewElement("w:titlePg");
            sectPr->InsertEndChild(titlePg);

            XMLElement* pgSz = doc.NewElement("w:pgSz");
            pgSz->SetAttribute("w:w", "11906");
            pgSz->SetAttribute("w:h", "16838");
            sectPr->InsertEndChild(pgSz);

            XMLElement* pgMar = doc.NewElement("w:pgMar");
            pgMar->SetAttribute("w:top", "1440");
            pgMar->SetAttribute("w:right", "1440");
            pgMar->SetAttribute("w:bottom", "1440");
            pgMar->SetAttribute("w:left", "1440");
            sectPr->InsertEndChild(pgMar);

            p->InsertEndChild(sectPr);
        }

        p->InsertEndChild(r);
        return p;
        };

    body->InsertFirstChild(createParagraph(to_utf8("«_»________2025 г."), true));
    body->InsertFirstChild(createParagraph(studentName, false));
    body->InsertFirstChild(createParagraph(group, false));

    body->InsertFirstChild(createParagraph(to_utf8("«_»________2025 г."), false));
    body->InsertFirstChild(createParagraph(to_utf8("/ И. Ю. Кулаков"), false));
    body->InsertFirstChild(createParagraph(to_utf8("Руководитель"), false));

    body->InsertFirstChild(createParagraph(to_utf8("ПТИ.ЛР 4092. 004.021"), false));
    body->InsertFirstChild(createParagraph(to_utf8("по специальности 09.03.01 Информатика и вычислительная техника"), false));
    body->InsertFirstChild(createParagraph(docType, false));
    body->InsertFirstChild(createParagraph(topic, false));

    body->InsertFirstChild(createParagraph(to_utf8("КАФЕДРА ИНФОРМАЦИОННЫХ ТЕХНОЛОГИЙ И СИСТЕМ"), false));
    body->InsertFirstChild(createParagraph(to_utf8("ПОЛИТЕХНИЧЕСКИЙ ИНСТИТУТ"), false));
    body->InsertFirstChild(createParagraph(to_utf8("«Новгородский государственный университет имени Ярослава Мудрого»"), false));
    body->InsertFirstChild(createParagraph(to_utf8("образовательное учреждение высшего образования"), false));
    body->InsertFirstChild(createParagraph(to_utf8("Федеральное государственное бюджетное"), false));
    body->InsertFirstChild(createParagraph(to_utf8("Министерство науки и высшего образования Российской Федерации"), false));

    if (doc.SaveFile(filePath.c_str()) != XML_SUCCESS) {
        std::cerr << "Ошибка при сохранении файла: " << filePath << std::endl;
    }
    else {
        std::cout << "ГОСТ-титульный лист успешно добавлен.\n";
    }
}