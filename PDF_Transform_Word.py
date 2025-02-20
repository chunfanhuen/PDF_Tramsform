import hashlib
import subprocess
from docx import Document
import os
import re
import requests
from concurrent.futures import ThreadPoolExecutor
import logging
import time
import json
from urllib.parse import quote

# 配置日志记录
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

CONFIG = {
    "translate_api": "microsoft",  # microsoft/baidu/deepl
    "microsoft_key": "YOUR_KEY",
    "microsoft_region": "YOUR_REGION",
    "baidu_appid": "YOUR_APPID",
    "baidu_secret": "YOUR_SECRET",
    "proxy": None,  # 例如："http://127.0.0.1:1080"
    "max_workers": 4,
    "request_timeout": 15,
}


def split_text(text, max_length=4500):
    """文本分割函数"""
    paragraphs = []
    current_para = []
    current_length = 0

    # 按句子分割，保留格式
    sentences = re.split(r'(?<=\n)|(?<=。|\.|\!|\?|；|;)\s*', text)

    for sentence in sentences:
        sentence = sentence.strip()
        if not sentence:
            continue

        sentence_length = len(sentence)
        if current_length + sentence_length > max_length and current_para:
            paragraphs.append('\n'.join(current_para))
            current_para = [sentence]
            current_length = sentence_length
        else:
            current_para.append(sentence)
            current_length += sentence_length + 1

    if current_para:
        paragraphs.append('\n'.join(current_para))
    return paragraphs


def translate_text(text, src_lang='en', target_lang='zh-CN', max_retries=3):
    """翻译函数"""
    if not text.strip():
        return ""

    for attempt in range(max_retries):
        try:
            if CONFIG["translate_api"] == "microsoft":
                return translate_microsoft(text, src_lang, target_lang)
            elif CONFIG["translate_api"] == "baidu":
                return translate_baidu(text, src_lang, target_lang)
            else:
                raise ValueError("不支持的翻译API")

        except Exception as e:
            logging.warning(f"翻译尝试 {attempt + 1} 失败: {str(e)}")
            time.sleep(2 ** attempt)
    return "[翻译失败]"


def translate_microsoft(text, src_lang, target_lang):
    """使用微软翻译API"""
    endpoint = "https://api.cognitive.microsofttranslator.com/translate"
    params = {
        "api-version": "3.0",
        "from": src_lang,
        "to": [target_lang]
    }
    headers = {
        "Ocp-Apim-Subscription-Key": CONFIG["microsoft_key"],
        "Ocp-Apim-Subscription-Region": CONFIG["microsoft_region"],
        "Content-Type": "application/json"
    }

    body = [{"text": text}]

    session = requests.Session()
    if CONFIG["proxy"]:
        session.proxies = {"https": CONFIG["proxy"]}

    response = session.post(
        endpoint,
        params=params,
        headers=headers,
        json=body,
        timeout=CONFIG["request_timeout"]
    )

    if response.status_code != 200:
        raise Exception(f"API错误: {response.text}")

    return response.json()[0]["translations"][0]["text"]


def translate_baidu(text, src_lang, target_lang):
    """使用百度翻译API"""
    url = "https://fanyi-api.baidu.com/api/trans/vip/translate"
    salt = str(time.time())
    sign_str = CONFIG["baidu_appid"] + text + salt + CONFIG["baidu_secret"]
    sign = hashlib.md5(sign_str.encode()).hexdigest()

    params = {
        "q": text,
        "from": src_lang,
        "to": target_lang,
        "appid": CONFIG["baidu_appid"],
        "salt": salt,
        "sign": sign
    }

    session = requests.Session()
    if CONFIG["proxy"]:
        session.proxies = {"https": CONFIG["proxy"]}

    response = session.get(
        url,
        params=params,
        timeout=CONFIG["request_timeout"]
    )

    result = response.json()
    if "error_code" in result:
        raise Exception(f"API错误: {result}")

    return "\n".join([item["dst"] for item in result["trans_result"]])
def pdf_to_docx_ocr_translate(pdf_path, docx_path, ocr_lang="eng", target_lang="zh-CN",
                              dpi=300, poppler_path=None, tesseract_path=None):
    """
    PDF转Word文档函数
    """
    # OCR语言映射（扩展更多语言支持）
    lang_map = {
        "eng": "en",
        "chi_sim": "zh-CN",
        "jpn": "ja",
        "deu": "de",
        "fra": "fr",
        "spa": "es"
    }

    src_lang = lang_map.get(ocr_lang.lower(), ocr_lang.split('_')[0])

    # 配置路径
    env = os.environ.copy()
    if poppler_path:
        env["PATH"] = f"{poppler_path}{os.pathsep}{env['PATH']}"

    tesseract_cmd = tesseract_path if tesseract_path else "tesseract"
    # 创建临时目录
    temp_dir = "temp_ocr_images"
    os.makedirs(temp_dir, exist_ok=True)

    try:
        # 修改 PDF 转图片命令（添加 -rx 300 -ry 300 参数）
        poppler_command = [
            "pdftoppm",
            "-r", str(dpi),
            "-rx", "300",  # 添加水平分辨率
            "-ry", "300",  # 添加垂直分辨率
            "-png",
            pdf_path,
            os.path.join(temp_dir, "page")
        ]

        result = subprocess.run(
            poppler_command,
            capture_output=True,
            text=True,
            env=env,
            timeout=60
        )

        if result.returncode != 0:
            logging.error(f"PDF转图片失败: {result.stderr}")
            return

        # 获取生成的图片列表
        image_list = sorted([
            os.path.abspath(os.path.join(temp_dir, f))
            for f in os.listdir(temp_dir)
            if f.endswith(".png")
        ], key=lambda x: int(re.search(r'page-(\d+)', x).group(1)))

        if not image_list:
            logging.error("未生成任何图片文件")
            return

        logging.info(f"成功生成 {len(image_list)} 张图片")

        # 创建Word文档
        doc = Document()
        doc.add_heading("OCR翻译文档", 0)

        # 使用线程池加速OCR处理
        with ThreadPoolExecutor(max_workers=CONFIG["max_workers"]) as executor:
            futures = []
            for img_path in image_list:
                future = executor.submit(
                    process_image,
                    img_path,
                    tesseract_cmd,
                    ocr_lang,
                    src_lang,
                    target_lang
                )
                futures.append((img_path, future))

            for img_path, future in futures:
                try:
                    ocr_text, translated_text = future.result()
                    add_to_document(doc, img_path, ocr_text, translated_text)
                except Exception as e:
                    logging.error(f"处理 {img_path} 失败: {str(e)}")
                    doc.add_paragraph(f"处理失败: {img_path}")

        # 保存文档
        doc.save(docx_path)
        logging.info(f"文档已保存至: {docx_path}")

    finally:
        # 清理临时文件
        for img_path in image_list:
            try:
                os.remove(img_path)
            except Exception as e:
                logging.warning(f"删除临时文件失败 {img_path}: {str(e)}")
        try:
            os.rmdir(temp_dir)
        except Exception as e:
            logging.warning(f"删除临时目录失败: {str(e)}")

    # 在执行 OCR 前添加验证
    if not os.path.exists(img_path):
        raise FileNotFoundError(f"图片文件不存在: {img_path}")

    #    在调用 subprocess 前添加路径打印
    logging.info(f"正在处理图片路径: {img_path}")


def process_image(img_path, tesseract_cmd, ocr_lang, src_lang, target_lang):
    """处理单张图片的OCR和翻译"""
    try:
        # OCR识别
        cmd = [
            tesseract_cmd,
            img_path,
            "stdout",
            "-l", ocr_lang,
            "--psm", "6",
            "--oem", "3",
            "-c", "preserve_interword_spaces=1",
            "-c", "tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789.,;:!?()%-'\""
        ]

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=30
        )

        ocr_text = result.stdout.strip() if result.returncode == 0 else ""
        logging.info(f"OCR完成: {img_path} [长度: {len(ocr_text)}]")

        if not ocr_text:
            return "[OCR无内容]", "[无需翻译]"

        # 翻译文本
        translated_text = translate_text(ocr_text, src_lang, target_lang)
        return ocr_text, translated_text

    except Exception as e:
        raise RuntimeError(f"处理图片时出错: {str(e)}") from e


def add_to_document(doc, img_path, ocr_text, translated_text):
    """将内容添加到Word文档"""
    doc.add_heading(os.path.basename(img_path), level=2)

    # 添加原文段落
    doc.add_paragraph("原文内容：")
    para = doc.add_paragraph(ocr_text)
    para.style = "Body Text"

    # 添加翻译段落
    doc.add_paragraph("翻译内容：")
    para = doc.add_paragraph(translated_text)
    para.style = "Body Text"

    # 添加分页符
    doc.add_page_break()


if __name__ == "__main__":
    # 配置路径（根据实际情况修改）
    input_pdf = r"./Ethereum_Whitepaper_-_Buterin_2014.pdf"
    output_docx = r"./Ethereum_Whitepaper_Translated.docx"
    poppler_path = r"D:/poppler-24.08.0/Library/bin"
    tesseract_path = r"D:/Tesseract-OCR/tesseract.exe"

    # 执行转换
    pdf_to_docx_ocr_translate(
        pdf_path=input_pdf,
        docx_path=output_docx,
        ocr_lang="eng",  # 中文文档使用 "chi_sim"
        target_lang="zh-CN",
        dpi=300,
        poppler_path=poppler_path,
        tesseract_path=tesseract_path
    )
