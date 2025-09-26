import os
import sys
import wmi
import ctypes
import subprocess
from collections import defaultdict
import time
import shutil
from tqdm import tqdm 

# --- 检查并获取管理员权限 ---
def is_admin():
    """检查脚本是否以管理员身份运行。"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin(command):
    """尝试以管理员身份重新启动脚本。"""
    try:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    except Exception as e:
        print(f"无法以管理员身份重新启动脚本: {e}")
        sys.exit(1)

# --- 全局进度条复制函数 (Tqdm) ---
def tqdm_copy_recursive(src_dir, dst_dir):
    """
    递归复制文件，并使用 tqdm 实现基于总字节数的全局进度条。
    """
    total_size = 0
    # 第一步：递归遍历，预先计算所有文件的总字节数
    print("  -> 正在预计算文件总大小，请稍候...")
    for dirpath, dirnames, filenames in os.walk(src_dir):
        for f in filenames:
            src_file = os.path.join(dirpath, f)
            try:
                # 累加文件大小，排除符号链接
                if not os.path.islink(src_file):
                    total_size += os.path.getsize(src_file)
            except Exception:
                pass

    # 第二步：使用 tqdm 包装整个复制过程
    print("  -> 开始复制文件...")
    with tqdm(total=total_size, unit='B', unit_scale=True, desc="📀 制作进度") as pbar:
        for dirpath, dirnames, filenames in os.walk(src_dir):
            # 构建目标路径
            relative_path = os.path.relpath(dirpath, src_dir)
            dest_dir = os.path.join(dst_dir, relative_path)
            
            # 创建目标目录
            os.makedirs(dest_dir, exist_ok=True)

            for filename in filenames:
                src_file = os.path.join(dirpath, filename)
                dest_file = os.path.join(dest_dir, filename)
                
                # 检查文件是否需要复制
                try:
                    if os.path.exists(dest_file) and os.path.getsize(src_file) == os.path.getsize(dest_file):
                        pbar.update(os.path.getsize(src_file))
                        continue
                except Exception:
                    pass
                        
                # 逐块复制文件并更新进度条
                try:
                    with open(src_file, 'rb') as fsrc, open(dest_file, 'wb') as fdst:
                        chunk_size = 65536  # 64KB
                        while True:
                            chunk = fsrc.read(chunk_size)
                            if not chunk:
                                break
                            fdst.write(chunk)
                            # 关键：更新全局进度条
                            pbar.update(len(chunk))
                    
                    # 复制文件状态/权限
                    shutil.copystat(src_file, dest_file)
                except Exception as e:
                    # 打印错误信息但不中断全局进度
                    pbar.write(f"\n[❌ 错误] 文件复制失败: {os.path.basename(src_file)} -> {e}")
                    # 在遇到错误时，更新进度条以避免卡住
                    try:
                        pbar.update(os.path.getsize(src_file))
                    except Exception:
                        pass
                    pass 
    print("  -> 文件复制操作完成。")


# --- 获取外部磁盘的函数 ---
def get_external_disks():
    """通过 WMI 筛选出所有外接的物理磁盘，并将分区信息聚合。"""
    external_disks = []
    try:
        c = wmi.WMI()
    except wmi.x_wmi:
        print("无法连接到WMI服务。请检查脚本是否以管理员身份运行。")
        return []

    internal_types = ["SATA", "IDE", "NVME"]
    
    for disk in c.Win32_DiskDrive():
        is_internal = False
        
        if disk.InterfaceType and disk.InterfaceType.upper() in internal_types:
            is_internal = True
        if is_internal:
            continue
        
        if "VEN" in disk.PNPDeviceID.upper() or "VID" in disk.PNPDeviceID.upper():
            pass
        else:
            continue

        disk_info = {
            "model": disk.Model,
            "size": int(disk.Size) if disk.Size else 0,
            "partitions": [],
            "device_id": disk.DeviceID # \\.\PHYSICALDRIVEx
        }
        for partition in disk.associators("Win32_DiskDriveToDiskPartition"):
            for logical_disk in partition.associators("Win32_LogicalDiskToPartition"):
                if logical_disk.DriveType == 3: # 确保是本地磁盘
                    disk_info["partitions"].append({
                        "device_id": logical_disk.DeviceID, # 盘符 (如 C:)
                        "volume_name": logical_disk.VolumeName,
                        "file_system": logical_disk.FileSystem,
                        "size": int(logical_disk.Size) if logical_disk.Size else 0
                    })
        external_disks.append(disk_info)
    return external_disks

# --- ISO 盘符获取函数 ---
def get_iso_drive_letter(iso_path):
    """通过查找常见的 ISO 特征（例如 'sources' 文件夹），来确定已挂载 ISO 文件的盘符。"""
    print("正在尝试通过文件内容定位ISO盘符...")
    for letter in 'CDEFGHIJKLMNOPQRSTUVWXYZ':
        drive = f'{letter}:\\'
        
        if drive.upper() in ["C:\\", os.environ.get("HOMEDRIVE", "").upper()]:
            continue
            
        if os.path.exists(drive):
            if os.path.isdir(os.path.join(drive, 'sources')):
                print(f"成功定位到ISO盘符: {drive}")
                return drive
    
    print("未能通过文件内容定位到ISO盘符。")
    return None

# --- DiskPart 自动化函数 ---
def make_usb_bootable(disk_index, iso_path, boot_mode="UEFI"):
    """使用 DiskPart 创建可启动U盘并复制文件。"""
    target_usb_mount_point = "Z:\\"
    script_path = os.path.join(os.environ['TEMP'], 'diskpart_script.txt')
    
    # 写入 DiskPart 脚本
    with open(script_path, 'w') as f:
        f.write(f'select disk {disk_index}\n')
        f.write('clean\n')
        
        if boot_mode == "UEFI":
            f.write('convert gpt\n')
            f.write('create partition primary\n')
            f.write('format fs=fat32 quick\n')
        else: # LEGACY
            f.write('create partition primary\n')
            f.write('select partition 1\n')
            f.write('active\n')
            f.write('format fs=ntfs quick\n')
        
        f.write(f'assign letter=Z\n')
        f.write('exit\n')

    print("正在执行 DiskPart 命令，请稍候...")
    
    # 执行 DiskPart 脚本
    try:
        subprocess.run(['diskpart', '/s', script_path], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
        print("DiskPart 脚本执行成功。")
    except subprocess.CalledProcessError as e:
        print(f"DiskPart 脚本执行失败: {e}")
        return False
    finally: # 修复了之前的 IndentationError 风险
        if os.path.exists(script_path):
            os.remove(script_path)

    time.sleep(10) 
    
    # 复制文件
    print("正在挂载ISO文件并复制内容到U盘...")
    
    # 挂载ISO文件到虚拟驱动器
    try:
        subprocess.run(['powershell', 'Mount-DiskImage', '-ImagePath', iso_path], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
    except subprocess.CalledProcessError as e:
        print(f"PowerShell 挂载ISO失败: {e}")
        return False
        
    try:
        # 获取ISO的挂载点
        iso_drive = get_iso_drive_letter(iso_path)
        if not iso_drive:
            print("无法找到挂载的ISO文件盘符。")
            return False

        # 检查目标盘符 Z:\
        if not os.path.exists(target_usb_mount_point):
            print("目标盘符 Z:\\ 尚未出现，等待 5 秒...")
            time.sleep(5) 
            if not os.path.exists(target_usb_mount_point):
                 print("目标盘符仍未出现，复制失败。")
                 return False

        os.makedirs(target_usb_mount_point, exist_ok=True)

        # 核心：调用全局进度条复制函数
        print(f"开始从 {iso_drive} 复制到 {target_usb_mount_point}...")
        tqdm_copy_recursive(iso_drive, target_usb_mount_point)

        print("\n文件复制完成。")
        return True
    except Exception as e:
        print(f"文件复制失败: {e}")
        return False
    finally:
        # 无论成功失败，都尝试卸载ISO文件
        try:
            subprocess.run(['powershell', 'Dismount-DiskImage', '-ImagePath', iso_path], creationflags=subprocess.CREATE_NO_WINDOW)
        except:
            pass

# --- 主程序 ---
def main():
    # 1. 权限检查
    if not is_admin():
        print("脚本正在以普通用户权限运行，将尝试重新启动为管理员。")
        run_as_admin(sys.argv)
        return

    iso_path = input("请输入ISO文件的完整路径：")
    if not os.path.exists(iso_path):
        print("文件不存在，请检查路径。")
        sys.exit(1)

    drives = get_external_disks()
    if not drives:
        print("未找到任何外接磁盘。")
        sys.exit(1)

    # 2. 磁盘选择界面
    print("找到以下外接磁盘，请选择：")
    for i, drive in enumerate(drives):
        disk_index = int(drive['device_id'].split('\\')[-1].replace('PHYSICALDRIVE', ''))
        
        print(f"{i}: {drive['model']} ({round(drive['size'] / (1024**3))} GB)")
        print(f"   - 物理设备ID: {drive['device_id']} (索引: {disk_index})")
        
        if not drive['partitions']:
            print("   - 未分区或无法获取分区信息")
            continue
        for part in drive["partitions"]:
            print(f"   - 分区: {part['device_id']} ({round(part['size'] / (1024**3))} GB)")
            print(f"     文件系统: {part['file_system'] if part['file_system'] else '未知'}")
            print(f"     卷标: {part['volume_name'] if part['volume_name'] else '无'}")
    
    # 3. 用户输入与确认
    try:
        choice = int(input("请输入数字选择磁盘："))
        if 0 <= choice < len(drives):
            selected_drive = drives[choice]
            disk_index = int(selected_drive['device_id'].split('\\')[-1].replace('PHYSICALDRIVE', ''))
            
            boot_mode = input("请选择启动模式 (输入 'UEFI' 或 'LEGACY'，默认UEFI): ").upper()
            if boot_mode not in ["UEFI", "LEGACY"]:
                 boot_mode = "UEFI" 
            
            print(f"\n警告：此操作将彻底清除磁盘 {selected_drive['model']} 的所有数据并制作 {boot_mode} 启动盘！")
            
            confirm = input("请再次确认，输入 'YES' 继续：")
            if confirm.upper() == 'YES':
                # 4. 执行制作启动盘
                if make_usb_bootable(disk_index, iso_path, boot_mode):
                    print("\n========== 成功 ==========")
                    print(f"启动U盘制作完成。请从BIOS/UEFI中选择 {boot_mode} 模式启动。")
                    print("==========================")
                else:
                    print("\n========== 失败 ==========")
                    print("启动U盘制作失败。请根据上面的错误信息进行检查。")
                    print("==========================")
            else:
                print("操作已取消。")
        else:
            print("无效的选择。")
    except ValueError:
        print("请输入一个有效的数字。")
    except Exception as e:
        print(f"发生致命错误: {e}")

    # 保持窗口打开，以便查看输出
    input("\n**操作完成。请按 Enter 键退出窗口...**")

if __name__ == "__main__":
    main()