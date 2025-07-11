import os
import time
import win32com.client

def export_single_slide_to_video(prs, slide_index, output_wmv):
    """导出单个幻灯片为视频"""
    powerpoint = prs.Application
    single_prs = powerpoint.Presentations.Add()
    prs.Slides(slide_index).Copy()
    single_prs.Slides.Paste()
    temp_pptx = os.path.abspath(f"temp_slide_{slide_index}.pptx")
    single_prs.SaveAs(temp_pptx)
    single_prs.CreateVideo(output_wmv, False, 5, 720, 30, 100)
    while single_prs.CreateVideoStatus != 3:
        print(f"幻灯片 {slide_index} 转换状态: {single_prs.CreateVideoStatus}")
        time.sleep(1)
    single_prs.Close()
    if os.path.exists(temp_pptx):
        os.remove(temp_pptx)

def convert_ppt_to_videos(src_pptx):
    """转换PPT为视频"""
    # 根据pptx文件名创建输出目录
    pptx_name = os.path.splitext(os.path.basename(src_pptx))[0]
    output_dir = pptx_name
    os.makedirs(output_dir, exist_ok=True)
    
    print(f"输入文件: {src_pptx}")
    print(f"输出目录: {output_dir}")
    
    powerpoint = None
    prs = None
    
    try:
        powerpoint = win32com.client.Dispatch('PowerPoint.Application.16')
        powerpoint.Visible = 1
        prs = powerpoint.Presentations.Open(src_pptx, WithWindow=False)
        slide_count = prs.Slides.Count
        
        print(f"开始转换，共 {slide_count} 页幻灯片")
        
        success_count = 0
        for i in range(1, slide_count + 1):
            wmv_path = os.path.abspath(os.path.join(output_dir, f"{pptx_name}_{i}.wmv"))
            print(f"正在导出第 {i} 页为视频: {wmv_path}")
            
            try:
                export_single_slide_to_video(prs, i, wmv_path)
                success_count += 1
                print(f"第 {i} 页导出完成")
            except Exception as e:
                print(f"第 {i} 页导出失败: {e}")
        
        print(f"\n转换完成！成功导出 {success_count}/{slide_count} 个视频到目录: {output_dir}")
        
    except Exception as e:
        print(f"转换过程中出现错误: {e}")
    finally:
        # 清理资源
        if prs:
            prs.Close()
        if powerpoint:
            powerpoint.Quit()

def main():
    """主函数 - 支持重复使用"""
    while True:
        print("\n" + "="*50)
        print("PPT转视频工具")
        print("="*50)
        
        # 获取输入文件
        src_pptx = input("请输入PPT文件路径（或输入 'quit' 退出）: ").strip().strip('"')
        
        if src_pptx.lower() == 'quit':
            print("程序退出")
            break
            
        if not src_pptx:
            print("请输入有效的文件路径")
            continue
            
        # 检查文件是否存在
        if not os.path.exists(src_pptx):
            print(f"文件不存在: {src_pptx}")
            continue
            
        # 检查文件扩展名
        if not src_pptx.lower().endswith('.pptx'):
            print("请选择.pptx文件")
            continue
            
        # 转换绝对路径
        src_pptx = os.path.abspath(src_pptx)
        
        # 开始转换
        convert_ppt_to_videos(src_pptx)
        
        # 询问是否继续
        while True:
            continue_choice = input("\n是否继续转换其他文件？(y/n): ").strip().lower()
            if continue_choice in ['y', 'yes', '是']:
                break
            elif continue_choice in ['n', 'no', '否']:
                print("程序退出")
                return
            else:
                print("请输入 y 或 n")

if __name__ == "__main__":
    main()