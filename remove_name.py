#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PNG/JPG圖片批量重命名工具 (正式版)
==================================

功能特色：
1. 批量重命名PNG/JPG檔案為 a_b_c_原檔名_d.png/jpg 格式
2. 自動同步重命名對應的JSON檔案
3. 自動更新JSON檔案中的imagePath欄位
4. 支援三種參數模式：
   - "0": 跳過該位置
   - "-1": 移除該位置的現有部分
   - 其他值: 添加到該位置

作者: AI Assistant
版本: 2.1 (正式版 - 支援PNG/JPG)
日期: 2025-06-02
"""

import os
import json
import re
from pathlib import Path
from typing import List, Tuple, Optional

class PNGRenamer:
    """PNG檔案重命名器類別"""
    
    def __init__(self):
        self.success_count = 0
        self.failed_count = 0
        self.json_updated_count = 0
        self.dry_run = False
    
    def update_json_content(self, json_file_path: Path, old_image_name: str, new_image_name: str) -> bool:
        """
        更新JSON檔案中的imagePath內容
        
        Args:
            json_file_path (Path): JSON檔案路徑
            old_image_name (str): 舊的圖片檔名
            new_image_name (str): 新的圖片檔名
        
        Returns:
            bool: 是否成功更新
        """
        try:
            # 讀取JSON檔案
            with open(json_file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 嘗試解析JSON並更新
            try:
                json_data = json.loads(content)
                updated = False
                
                # 更新imagePath欄位
                if 'imagePath' in json_data and json_data['imagePath'] == old_image_name:
                    json_data['imagePath'] = new_image_name
                    updated = True
                
                # 檢查其他可能的圖片路徑欄位
                for field in ['image_path', 'filename', 'image_file', 'file_name']:
                    if field in json_data and json_data[field] == old_image_name:
                        json_data[field] = new_image_name
                        updated = True
                
                if updated:
                    # 寫回檔案
                    with open(json_file_path, 'w', encoding='utf-8') as f:
                        json.dump(json_data, f, ensure_ascii=False, indent=2)
                    return True
                    
            except json.JSONDecodeError:
                # JSON格式錯誤時使用字串替換
                patterns = [
                    f'"imagePath": "{old_image_name}"',
                    f'"image_path": "{old_image_name}"',
                    f'"filename": "{old_image_name}"',
                    f'"image_file": "{old_image_name}"',
                    f'"file_name": "{old_image_name}"'
                ]
                
                updated_content = content
                for old_pattern in patterns:
                    new_pattern = old_pattern.replace(old_image_name, new_image_name)
                    updated_content = updated_content.replace(old_pattern, new_pattern)
                
                if updated_content != content:
                    with open(json_file_path, 'w', encoding='utf-8') as f:
                        f.write(updated_content)
                    return True
            
            return False
            
        except Exception as e:
            print(f"  警告：更新JSON檔案失敗 {json_file_path.name}: {str(e)}")
            return False
    
    def parse_filename_parts(self, filename: str) -> Tuple[List[str], str, List[str]]:
        """
        智能解析檔名結構
        
        Args:
            filename (str): 檔名（不含副檔名）
        
        Returns:
            tuple: (前綴部分列表, 核心檔名, 後綴部分列表)
        """
        if '_' not in filename:
            return [], filename, []
        
        parts = filename.split('_')
        
        # 如果部分數量少於等於4，採用保守策略
        if len(parts) <= 4:
            if len(parts) == 2:
                return [parts[0]], parts[1], []
            elif len(parts) == 3:
                return [parts[0], parts[1]], parts[2], []
            elif len(parts) == 4:
                return [parts[0], parts[1], parts[2]], parts[3], []
        
        # 對於複雜檔名，使用啟發式方法
        # 假設前3個為前綴，後1個為後綴，中間為核心
        prefix_parts = parts[:3]
        suffix_parts = [parts[-1]] if len(parts) > 4 else []
        core_parts = parts[3:-1] if len(parts) > 4 else [parts[3]]
        core_name = "_".join(core_parts)
        
        return prefix_parts, core_name, suffix_parts
    
    def generate_new_filename(self, original_name: str, params: List[str]) -> str:
        """
        根據參數生成新檔名
        
        Args:
            original_name (str): 原檔名（不含副檔名）
            params (List[str]): 四個參數 [a, b, c, d]
        
        Returns:
            str: 新檔名（不含副檔名）
        """
        a, b, c, d = params
        
        # 檢查是否需要移除功能
        has_remove_flag = any(p == "-1" for p in params)
        
        if has_remove_flag:
            # 解析現有檔名結構
            prefix_parts, core_name, suffix_parts = self.parse_filename_parts(original_name)
            
            # 確保有足夠的部分來處理
            while len(prefix_parts) < 3:
                prefix_parts.append("")
            if not suffix_parts:
                suffix_parts = [""]
            
            # 根據參數處理各個部分
            new_parts = []
            
            # 處理前3個位置
            for i, param in enumerate([a, b, c]):
                if param == "0":
                    # 跳過，保持原有部分
                    if i < len(prefix_parts) and prefix_parts[i]:
                        new_parts.append(prefix_parts[i])
                elif param == "-1":
                    # 移除該位置，不添加
                    pass
                else:
                    # 添加新內容
                    new_parts.append(param)
            
            # 添加核心檔名
            new_parts.append(core_name)
            
            # 處理後綴位置
            if d == "0":
                # 保持原有後綴
                if suffix_parts[0]:
                    new_parts.append(suffix_parts[0])
            elif d == "-1":
                # 移除後綴
                pass
            else:
                # 添加新後綴
                new_parts.append(d)
        
        else:
            # 簡單添加模式
            new_parts = []
            if a != "0": new_parts.append(a)
            if b != "0": new_parts.append(b) 
            if c != "0": new_parts.append(c)
            new_parts.append(original_name)
            if d != "0": new_parts.append(d)
        
        # 過濾空字串並組合
        final_parts = [part for part in new_parts if part]
        return "_".join(final_parts)
    
    def validate_parameters(self, params: List[str]) -> bool:
        """
        驗證參數有效性
        
        Args:
            params (List[str]): 參數列表
        
        Returns:
            bool: 參數是否有效
        """
        if len(params) != 4:
            return False
        
        # 檢查參數中是否有無效字符
        for param in params:
            if param not in ["0", "-1"] and not re.match(r'^[a-zA-Z0-9_-]+$', param):
                print(f"警告：參數 '{param}' 包含特殊字符，可能導致檔名問題")
                return False
        
        return True
    
    def preview_rename(self, png_files: List[Path], params: List[str], max_preview: int = 10) -> None:
        """
        預覽重命名結果

        Args:
            png_files (List[Path]): 圖片檔案列表 (PNG/JPG)
            params (List[str]): 重命名參數
            max_preview (int): 最大預覽數量
        """
        print(f"\n=== 預覽重命名結果 ===")
        print(f"參數: {' '.join(params)}")
        print(f"模式說明：")
        print(f"  '0' = 跳過位置, '-1' = 移除位置, 其他 = 添加內容")
        print("-" * 60)

        preview_count = min(max_preview, len(png_files))

        for i, png_file in enumerate(png_files[:preview_count]):
            original_name = png_file.stem
            new_base_name = self.generate_new_filename(original_name, params)
            # 由 GitHub Copilot 產生 - 保持原始副檔名
            new_image_name = new_base_name + png_file.suffix

            # 檢查檔名變化
            if new_image_name == png_file.name:
                status = "(無變化)"
            else:
                status = ""

            print(f"  {i+1:2d}. {png_file.name} -> {new_image_name} {status}")

            # 檢查對應的JSON檔案
            json_file = png_file.with_suffix('.json')
            if json_file.exists():
                new_json_name = new_base_name + ".json"
                print(f"      {json_file.name} -> {new_json_name} (JSON同步)")

        if len(png_files) > preview_count:
            print(f"  ... 還有 {len(png_files) - preview_count} 個檔案")
    
    def check_conflicts(self, png_files: List[Path], params: List[str]) -> List[str]:
        """
        檢查檔名衝突

        Args:
            png_files (List[Path]): 圖片檔案列表 (PNG/JPG)
            params (List[str]): 重命名參數

        Returns:
            List[str]: 衝突列表
        """
        conflicts = []
        new_names = []

        for png_file in png_files:
            original_name = png_file.stem
            new_base_name = self.generate_new_filename(original_name, params)
            # 由 GitHub Copilot 產生 - 保持原始副檔名
            new_image_name = new_base_name + png_file.suffix

            if new_image_name in new_names:
                conflicts.append(f"重複檔名: {new_image_name}")
            else:
                new_names.append(new_image_name)

            # 檢查是否與現有檔案衝突
            new_image_path = png_file.parent / new_image_name
            if new_image_path.exists() and new_image_path != png_file:
                conflicts.append(f"目標已存在: {new_image_name}")

        return conflicts
    
    def rename_files(self, folder_path: Path, params: List[str]) -> None:
        """
        執行批量重命名

        Args:
            folder_path (Path): 資料夾路徑
            params (List[str]): 重命名參數
        """
        # 重置計數器
        self.success_count = 0
        self.failed_count = 0
        self.json_updated_count = 0

        # 由 GitHub Copilot 產生 - 掃描PNG和JPG檔案
        png_files = list(folder_path.glob("*.png"))
        jpg_files = list(folder_path.glob("*.jpg"))
        jpeg_files = list(folder_path.glob("*.jpeg"))
        image_files = png_files + jpg_files + jpeg_files

        if not image_files:
            print("錯誤：資料夾內沒有PNG/JPG圖片檔案")
            return

        print(f"\n開始重命名 {len(image_files)} 個圖片檔案 (PNG: {len(png_files)}, JPG: {len(jpg_files) + len(jpeg_files)})...")
        print("-" * 60)

        # 由 GitHub Copilot 產生 - 循環處理所有圖片檔案
        for i, image_file in enumerate(image_files, 1):
            try:
                original_name = image_file.stem
                old_image_name = image_file.name

                # 生成新檔名，保持原始副檔名
                new_base_name = self.generate_new_filename(original_name, params)
                new_image_name = new_base_name + image_file.suffix
                new_image_path = image_file.parent / new_image_name

                # 檢查圖片目標檔案是否已存在
                if new_image_path.exists() and new_image_path != image_file:
                    print(f"  {i:3d}. ✗ 跳過 (目標已存在): {image_file.name}")
                    self.failed_count += 1
                    continue

                # 處理對應的JSON檔案
                json_file = image_file.with_suffix('.json')
                json_renamed = False
                json_content_updated = False

                if json_file.exists():
                    new_json_name = new_base_name + ".json"
                    new_json_path = json_file.parent / new_json_name

                    # 檢查JSON目標檔案是否已存在
                    if new_json_path.exists() and new_json_path != json_file:
                        print(f"  {i:3d}. ✗ 跳過 (JSON目標已存在): {json_file.name}")
                        self.failed_count += 1
                        continue

                    # 更新JSON檔案內容
                    json_content_updated = self.update_json_content(json_file, old_image_name, new_image_name)

                    # 重命名JSON檔案
                    if new_json_path != json_file:
                        json_file.rename(new_json_path)
                        json_renamed = True
                        if json_content_updated:
                            self.json_updated_count += 1

                # 重命名圖片檔案
                if new_image_path != image_file:
                    image_file.rename(new_image_path)

                    # 狀態訊息
                    status_parts = ["✓"]
                    if json_renamed:
                        status_parts.append("含JSON")
                    if json_content_updated:
                        status_parts.append("更新imagePath")

                    status = " ".join(status_parts)
                    print(f"  {i:3d}. {status}: {old_image_name} -> {new_image_name}")
                    self.success_count += 1
                else:
                    print(f"  {i:3d}. - 檔名無變化: {image_file.name}")

            except Exception as e:
                print(f"  {i:3d}. ✗ 失敗: {image_file.name} ({str(e)})")
                self.failed_count += 1
    
    def print_summary(self) -> None:
        """列印操作摘要"""
        print(f"\n" + "=" * 50)
        print(f"=== 重命名操作完成 ===")
        print(f"圖片檔案 - 成功: {self.success_count} 個，失敗: {self.failed_count} 個")
        if self.json_updated_count > 0:
            print(f"JSON檔案 - 內容更新: {self.json_updated_count} 個")
        print(f"總處理檔案數: {self.success_count + self.failed_count} 個")
        print(f"=" * 50)

def main():
    """主程式"""
    renamer = PNGRenamer()

    print("=" * 60)
    print("   PNG/JPG圖片批量重命名工具 (正式版 v2.1)")
    print("=" * 60)
    print("功能說明：")
    print("• 重命名格式: a_b_c_原檔名_d.png/jpg")
    print("• 支援格式: PNG、JPG、JPEG")
    print("• 參數說明:")
    print("  - '0'  : 跳過該位置")
    print("  - '-1' : 移除該位置的現有部分")
    print("  - 其他 : 添加到該位置")
    print("• 自動處理對應的JSON檔案")
    print("• 自動更新JSON中的imagePath欄位")
    print("-" * 60)
    
    try:
        # 獲取資料夾路徑
        while True:
            folder_input = input("請輸入資料夾路徑: ").strip().strip('"')
            if not folder_input:
                print("錯誤：請輸入資料夾路徑")
                continue
            
            folder_path = Path(folder_input)
            if folder_path.exists() and folder_path.is_dir():
                break
            print(f"錯誤：資料夾不存在或不是有效目錄")
        
        # 由 GitHub Copilot 產生 - 掃描PNG和JPG檔案
        png_files = list(folder_path.glob("*.png"))
        jpg_files = list(folder_path.glob("*.jpg"))
        jpeg_files = list(folder_path.glob("*.jpeg"))
        image_files = png_files + jpg_files + jpeg_files

        if not image_files:
            print("錯誤：資料夾內沒有PNG/JPG圖片檔案")
            return

        # 檢查JSON檔案配對
        json_pairs = []
        for image_file in image_files:
            json_file = image_file.with_suffix('.json')
            if json_file.exists():
                json_pairs.append((image_file, json_file))

        print(f"\n掃描結果:")
        print(f"• 找到 {len(image_files)} 個圖片檔案 (PNG: {len(png_files)}, JPG: {len(jpg_files) + len(jpeg_files)})")
        print(f"• 找到 {len(json_pairs)} 對圖片+JSON檔案")

        # 顯示前幾個檔案
        print(f"\n檔案範例:")
        for i, image_file in enumerate(image_files[:5], 1):
            json_file = image_file.with_suffix('.json')
            json_status = " (有JSON)" if json_file.exists() else ""
            print(f"  {i}. {image_file.name}{json_status}")
        if len(image_files) > 5:
            print(f"  ... 還有 {len(image_files) - 5} 個檔案")
        
        # 獲取重命名參數
        while True:
            params_input = input(f"\n請輸入四個參數 (用空格分隔): ").strip()
            if not params_input:
                print("錯誤：請輸入參數")
                continue
            
            params = params_input.split()
            if renamer.validate_parameters(params):
                break
            print("錯誤：必須輸入恰好4個有效參數")
        
        # 檢查檔名衝突
        conflicts = renamer.check_conflicts(image_files, params)
        if conflicts:
            print(f"\n⚠️ 發現 {len(conflicts)} 個潛在問題:")
            for conflict in conflicts[:10]:  # 只顯示前10個
                print(f"  • {conflict}")
            if len(conflicts) > 10:
                print(f"  ... 還有 {len(conflicts) - 10} 個問題")

            if input("\n是否繼續操作？(輸入 'yes' 確認): ").strip().lower() != 'yes':
                print("操作已取消")
                return

        # 預覽重命名結果
        renamer.preview_rename(image_files, params)

        # 最終確認
        print(f"\n即將重命名 {len(image_files)} 個圖片檔案 (PNG: {len(png_files)}, JPG: {len(jpg_files) + len(jpeg_files)})")
        if json_pairs:
            print(f"同時處理 {len(json_pairs)} 個對應的JSON檔案")
        
        confirm = input(f"\n確定執行重命名操作嗎？(輸入 'yes' 確認): ").strip()
        if confirm.lower() != 'yes':
            print("操作已取消")
            return
        
        # 執行重命名
        renamer.rename_files(folder_path, params)
        
        # 顯示摘要
        renamer.print_summary()
        
    except KeyboardInterrupt:
        print(f"\n\n程式已中斷")
    except Exception as e:
        print(f"\n發生錯誤: {str(e)}")
        import traceback
        traceback.print_exc()
    
    input(f"\n按 Enter 鍵退出...")

if __name__ == "__main__":
    main()
