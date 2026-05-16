import zipfile
import os

# 打开一个成图文件，查看图表XML结构
test_file = r"E:\杂七杂八\trea项目\xiang\dist\模板\成图模板.xlsx"

if os.path.exists(test_file):
    with zipfile.ZipFile(test_file, 'r') as z:
        if 'xl/charts/chart1.xml' in z.namelist():
            content = z.read('xl/charts/chart1.xml').decode('utf-8')
            # 保存到文件以便查看
            with open('chart_debug.xml', 'w', encoding='utf-8') as f:
                f.write(content)
            print("图表XML已保存到 chart_debug.xml")
            # 搜索标题相关的部分
            import re
            title_parts = re.findall(r'<c:title>.*?</c:title>', content, flags=re.DOTALL)
            print("\n找到的标题部分:")
            for i, part in enumerate(title_parts):
                print(f"\n=== 标题部分 {i+1} ===")
                print(part)
        else:
            print("未找到 xl/charts/chart1.xml")
else:
    print(f"文件不存在: {test_file}")
