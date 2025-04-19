import os  # 导入os模块，用于处理文件和目录路径
import tempfile  # 导入tempfile模块，用于创建临时文件和目录
import shutil  # 导入shutil模块，用于高级文件操作，如删除目录树
import win32com.client # 导入pywin32库，用于COM操作（如调用PowerPoint）
import pythoncom # 导入pythoncom，用于COM初始化/反初始化
from pypdf import PdfWriter, PdfReader  # 从pypdf库导入PdfWriter用于写入PDF，PdfReader用于读取PDF
from PIL import Image  # 从Pillow库导入Image模块，用于处理图片文件
from fpdf import FPDF  # 从fpdf库导入FPDF类，用于从图片创建PDF
from docx2pdf import convert  # 从docx2pdf库导入convert函数，用于将DOCX文件转换为PDF


def merge_files(file_list, output_filename):
    """将列表中的PDF、DOCX和图片文件合并成一个输出PDF文件。"""
    pdf_writer = PdfWriter()  # 创建一个PdfWriter对象，用于构建最终的合并PDF
    temp_dir = None  # 初始化临时目录变量为None，确保在finally块中可以安全检查

    # 检查输入的文件列表是否为空
    if not file_list:
        print("警告：没有选择任何文件进行合并。")  # 如果列表为空，打印警告信息
        return False, "没有选择任何文件。"  # 返回失败状态和消息

    try:
        # 创建一个唯一的临时目录来存放转换后的文件
        temp_dir = tempfile.mkdtemp()
        print(f"创建临时目录: {temp_dir}")  # 打印临时目录路径

        converted_files_for_cleanup = []  # 创建一个列表，用于记录需要清理的临时转换文件（虽然当前代码未显式使用此列表进行清理，但保留了结构）

        print(f"开始处理 {len(file_list)} 个文件:")  # 打印将要处理的文件总数
        # 遍历用户提供的文件列表
        for index, file_path in enumerate(file_list):
            filename = os.path.basename(file_path)  # 获取文件的基本名称（不含路径）
            print(f"  - ({index + 1}/{len(file_list)}) 正在处理: {filename}")  # 打印当前处理的文件名和进度
            _, ext = os.path.splitext(filename.lower())  # 获取文件的小写扩展名

            try:
                # --- 处理 PDF 文件 ---
                if ext == '.pdf':
                    pdf_reader = PdfReader(file_path)  # 创建PdfReader对象读取PDF文件
                    # 遍历PDF的每一页
                    for page in pdf_reader.pages:
                        pdf_writer.add_page(page)  # 将页面添加到PdfWriter对象中
                    print(f"    - 添加 PDF 页面: {len(pdf_reader.pages)} 页")  # 打印添加的页数

                # --- 处理 DOCX 文件 ---
                elif ext == '.docx':
                    # 构建临时PDF文件的完整路径
                    temp_pdf_path = os.path.join(temp_dir, f"{os.path.splitext(filename)[0]}.pdf")
                    print(f"    - 正在转换 Word 文件: {filename} -> {os.path.basename(temp_pdf_path)}") # 打印转换信息
                    convert(file_path, temp_pdf_path)  # 调用docx2pdf库进行转换
                    # 检查转换后的PDF文件是否存在
                    if os.path.exists(temp_pdf_path):
                        pdf_reader = PdfReader(temp_pdf_path)  # 读取转换后的临时PDF
                        # 遍历临时PDF的每一页
                        for page in pdf_reader.pages:
                            pdf_writer.add_page(page)  # 将页面添加到PdfWriter对象中
                        print(f"    - 添加转换后的 PDF 页面: {len(pdf_reader.pages)} 页") # 打印添加的页数
                        # converted_files_for_cleanup.append(temp_pdf_path) # 标记此临时文件以便后续清理（如果需要单独处理）
                    else:
                        print(f"警告：Word 文件转换失败: {filename}") # 如果转换失败，打印警告
                        # 可选：在这里抛出错误或决定是否继续处理其他文件

                # --- 处理 PowerPoint 文件 (.ppt, .pptx) ---
                elif ext in ['.ppt', '.pptx']:
                    powerpoint = None
                    presentation = None
                    temp_pdf_path = os.path.join(temp_dir, f"{os.path.splitext(filename)[0]}.pdf")
                    print(f"    - 正在转换 PowerPoint 文件: {filename} -> {os.path.basename(temp_pdf_path)}")
                    try:
                        pythoncom.CoInitialize() # 初始化COM环境
                        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
                        # powerpoint.Visible = 1 # 可选：使PowerPoint窗口可见
                        presentation = powerpoint.Presentations.Open(file_path, WithWindow=False)
                        # 使用 ppSaveAsPDF (常量值 32)
                        presentation.SaveAs(temp_pdf_path, 32)
                        print(f"    - PowerPoint 文件转换成功")
                    except Exception as ppt_e:
                        print(f"错误：转换 PowerPoint 文件 '{filename}' 时出错: {ppt_e}")
                    finally:
                        if presentation:
                            presentation.Close()
                        if powerpoint:
                            powerpoint.Quit()
                        pythoncom.CoUninitialize() # 反初始化COM环境

                    # 检查转换后的PDF文件是否存在并添加到writer
                    if os.path.exists(temp_pdf_path):
                        try:
                            pdf_reader = PdfReader(temp_pdf_path)
                            for page in pdf_reader.pages:
                                pdf_writer.add_page(page)
                            print(f"    - 添加转换后的 PDF 页面: {len(pdf_reader.pages)} 页")
                        except Exception as read_ppt_pdf_e:
                            print(f"错误：读取转换后的 PowerPoint PDF '{os.path.basename(temp_pdf_path)}' 时出错: {read_ppt_pdf_e}")
                    else:
                        print(f"警告：PowerPoint 文件转换失败或未生成PDF: {filename}")

                # --- 处理图片文件 --- 
                elif ext in ['.jpg', '.jpeg', '.png']:
                    # 构建临时PDF文件的完整路径
                    temp_pdf_path = os.path.join(temp_dir, f"{os.path.splitext(filename)[0]}.pdf")
                    print(f"    - 正在转换图片文件: {filename} -> {os.path.basename(temp_pdf_path)}") # 打印转换信息
                    try:
                        img = Image.open(file_path)  # 使用Pillow打开图片文件
                        img_width, img_height = img.size  # 获取图片的宽度和高度（像素）
                        # 注意：FPDF默认单位是毫米(mm)，但这里使用'pt'（点）作为单位，1点 = 1/72英寸
                        # A4 尺寸大约是 595 x 842 点
                        # 这里直接使用图片的像素尺寸作为PDF页面的尺寸（以点为单位），FPDF会处理单位转换
                        pdf = FPDF(unit="pt", format=(img_width, img_height)) # 创建FPDF对象，设置单位和页面尺寸
                        pdf.add_page()  # 添加一个页面
                        # 将图片添加到PDF页面，(0, 0)是左上角坐标，指定宽度和高度
                        pdf.image(file_path, 0, 0, img_width, img_height)
                        pdf.output(temp_pdf_path, "F")  # 将PDF内容输出（保存）到临时文件
                        img.close()  # 关闭图片文件，释放资源

                        # 检查转换后的PDF文件是否存在
                        if os.path.exists(temp_pdf_path):
                            pdf_reader = PdfReader(temp_pdf_path)  # 读取转换后的临时PDF
                            # 遍历临时PDF的每一页（对于图片转换，通常只有一页）
                            for page in pdf_reader.pages:
                                pdf_writer.add_page(page)  # 将页面添加到PdfWriter对象中
                            print(f"    - 添加转换后的 PDF 页面: 1 页") # 打印添加的页数
                            # converted_files_for_cleanup.append(temp_pdf_path) # 标记此临时文件
                        else:
                             print(f"警告：图片文件转换失败: {filename}") # 如果转换失败，打印警告

                    except Exception as img_e:
                        print(f"错误：处理图片文件 '{filename}' 时出错: {img_e}") # 打印图片处理过程中的具体错误
                        # 确保即使发生错误也尝试关闭图片文件（如果已打开）
                        if 'img' in locals() and img:
                            img.close()

                # --- 处理不支持的文件类型 ---
                else:
                    # 检查是否是旧版 .doc 文件，如果是，则提示不支持
                    if ext == '.doc':
                         print(f"警告：跳过不支持的旧版 Word 文件 (.doc): {filename}。请先转换为 .docx 格式。")
                    else:
                        print(f"警告：跳过不支持的文件类型: {filename}") # 如果文件类型不支持，打印警告

            # 捕获处理单个文件时可能发生的任何异常
            except Exception as file_e:
                print(f"错误：处理文件 '{filename}' 时出错: {file_e}") # 打印处理该文件时发生的错误
                # 在这里可以决定是继续处理下一个文件还是中断整个合并过程
                # return False, f"处理文件 '{filename}' 时出错: {file_e}" # 如果需要中断，可以取消这行的注释

        # --- 写入最终的合并PDF --- 
        # 检查是否有页面被添加到pdf_writer中
        if len(pdf_writer.pages) > 0:
            try:
                # 以二进制写入模式('wb')打开指定的输出文件
                with open(output_filename, 'wb') as out_pdf:
                    pdf_writer.write(out_pdf)  # 将pdf_writer中的所有页面写入输出文件
                print(f"\n成功！已将 {len(file_list)} 个文件（或其转换结果）合并到 '{output_filename}'") # 打印成功信息
                return True, f"成功合并 {len(file_list)} 个文件到 {output_filename}" # 返回成功状态和消息
            # 捕获写入文件时可能发生的异常
            except Exception as write_e:
                print(f"错误：无法写入合并后的 PDF 文件 '{output_filename}': {write_e}") # 打印写入错误
                return False, f"写入输出文件时出错: {write_e}" # 返回失败状态和消息
        else:
            # 如果没有任何页面被成功添加
            print("没有成功添加任何页面，未生成输出文件。")
            return False, "没有有效页面可合并。" # 返回失败状态和消息

    # 捕获整个合并过程中（如创建临时目录时）可能发生的任何未预料的异常
    except Exception as general_e:
        print(f"合并过程中发生意外错误: {general_e}") # 打印通用错误
        return False, f"合并过程中发生错误: {general_e}" # 返回失败状态和消息

    # --- 清理临时文件 --- 
    finally:
        # 无论合并成功与否，finally块中的代码总会执行
        # 检查临时目录是否已创建并且仍然存在
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)  # 递归删除临时目录及其所有内容
                print(f"已清理临时目录: {temp_dir}") # 打印清理成功信息
            # 捕获清理临时目录时可能发生的异常
            except Exception as clean_e:
                print(f"警告：无法完全清理临时目录 '{temp_dir}': {clean_e}") # 打印清理失败的警告

# 主程序入口块已被移除，因为现在合并操作由GUI触发。
# if __name__ == "__main__":
#     # 示例用法（如果需要单独测试此脚本）
#     test_files = ["path/to/your/file1.pdf", "path/to/your/image.png", "path/to/your/document.docx"]
#     output_file = "merged_test_output.pdf"
#     success, message = merge_files(test_files, output_file)
#     print(message)