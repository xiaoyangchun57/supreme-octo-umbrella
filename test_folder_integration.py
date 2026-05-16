import os
import sys
sys.path.insert(0, 'src')

from app.modules.folder_integration import FolderIntegrator

# 测试文件夹整合
def test_folder_integration():
    # 设置测试路径
    root_path = r'E:\杂七杂八\trea项目\xiang\dist\成果'
    
    print(f"测试目录: {root_path}")
    print(f"目录存在: {os.path.exists(root_path)}")
    
    if os.path.exists(root_path):
        print("\n目录内容:")
        for item in os.listdir(root_path):
            item_path = os.path.join(root_path, item)
            if os.path.isfile(item_path):
                print(f"  文件: {item}")
            else:
                print(f"  文件夹: {item}")
    
    # 检查对应表是否存在
    structure_path = os.path.join(root_path, '对应表.XLSX')
    print(f"\n对应表路径: {structure_path}")
    print(f"对应表存在: {os.path.exists(structure_path)}")
    
    structure_path_xlsx = os.path.join(root_path, '对应表.xlsx')
    print(f"对应表路径(xlsx): {structure_path_xlsx}")
    print(f"对应表存在(xlsx): {os.path.exists(structure_path_xlsx)}")
    
    structure_path_xlsm = os.path.join(root_path, '对应表.xlsm')
    print(f"对应表路径(xlsm): {structure_path_xlsm}")
    print(f"对应表存在(xlsm): {os.path.exists(structure_path_xlsm)}")
    
    # 运行整合
    template_dir = r'E:\杂七杂八\trea项目\xiang\dist\模板'
    print(f"\n模板目录: {template_dir}")
    print(f"模板目录存在: {os.path.exists(template_dir)}")
    
    integrator = FolderIntegrator(template_dir=template_dir)
    
    print("\n=== 开始整合 ===")
    results = integrator.process_all(root_path)
    
    print("\n=== 整合结果 ===")
    print(f"成功: {len(results['success'])}")
    print(f"失败: {len(results['failed'])}")
    
    # 输出日志
    print("\n=== 操作日志 ===")
    print(integrator.log_text)

if __name__ == '__main__':
    test_folder_integration()
