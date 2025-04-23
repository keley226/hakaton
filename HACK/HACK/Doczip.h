#include <iostream>
#include <filesystem>
#include <fstream>
#include <zip.h>
#include <tinyxml2.h>


namespace fs = std::filesystem;
using namespace tinyxml2;

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
void zipFolder(const std::string& folder, const std::string& outZip) {
    int errorp;
    zip_t* archive = zip_open(outZip.c_str(), ZIP_CREATE | ZIP_TRUNCATE, &errorp);
    if (!archive) {
        std::cerr << "Не удалось создать архив: " << zip_strerror(archive) << std::endl;
        return;
    }

    for (const auto& entry : fs::recursive_directory_iterator(folder)) {
        if (entry.is_directory()) continue;

        std::string relativePath = fs::relative(entry.path(), folder).string();
        std::replace(relativePath.begin(), relativePath.end(), '\\', '/');

        std::ifstream file(entry.path(), std::ios::binary | std::ios::ate);
        if (!file.is_open()) {
            std::cerr << "Ошибка открытия файла: " << entry.path() << std::endl;
            continue;
        }

        std::streamsize size = file.tellg();
        file.seekg(0, std::ios::beg);
        std::vector<char> buffer(size);
        if (!file.read(buffer.data(), size)) {
            std::cerr << "Ошибка чтения файла: " << entry.path() << std::endl;
            continue;
        }

        zip_source_t* source = zip_source_buffer_create(buffer.data(), buffer.size(), 0, nullptr);
        if (!source) {
            std::cerr << "Ошибка создания источника для: " << relativePath << std::endl;
            continue;
        }

        if (zip_file_add(archive, relativePath.c_str(), source, ZIP_FL_ENC_UTF_8 | ZIP_FL_OVERWRITE) < 0) {
            zip_source_free(source);
            std::cerr << "Ошибка добавления файла в архив: " << zip_strerror(archive) << std::endl;
            continue;
        }
    }

    if (zip_close(archive) < 0) {
        std::cerr << "Ошибка закрытия архива: " << zip_strerror(archive) << std::endl;
    }
}