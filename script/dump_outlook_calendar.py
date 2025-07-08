#!/usr/bin/env python3
"""
修正版完整時區感知Mac Outlook Calendar Reader
包含Calendar_UID和Record_ModDate欄位，修正CSV格式問題
"""

import sqlite3
import csv
import sys
import os
from datetime import datetime, timezone, timedelta
from pathlib import Path
import struct
import re
from html import unescape
import argparse

class CompleteFixedTimeZoneOutlookParser:
    def __init__(self, user_timezone='UTC+8'):
        self.outlook_data_path = os.path.expanduser("~/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data")
        self.db_path = os.path.join(self.outlook_data_path, "Outlook.sqlite")
        self.user_timezone = self.parse_timezone(user_timezone)
        
    def parse_timezone(self, tz_string):
        """解析時區字串"""
        if tz_string.upper() == 'UTC':
            return timezone.utc
        elif tz_string.upper().startswith('UTC'):
            sign = 1 if '+' in tz_string else -1
            try:
                offset_str = tz_string.split('+' if '+' in tz_string else '-')[1]
                if ':' in offset_str:
                    hours, minutes = map(int, offset_str.split(':'))
                else:
                    hours = int(offset_str)
                    minutes = 0
                offset = sign * (hours * 60 + minutes)
                return timezone(timedelta(minutes=offset))
            except:
                print(f"警告: 無法解析時區 '{tz_string}'，使用 UTC+8")
                return timezone(timedelta(hours=8))
        else:
            return timezone(timedelta(hours=8))
    
    def minutes_since_1601_to_datetime(self, minutes):
        """將從 1601-01-01 UTC 開始的分鐘數轉換為 UTC datetime"""
        try:
            filetime_epoch = datetime(1601, 1, 1, tzinfo=timezone.utc)
            dt_utc = filetime_epoch + timedelta(minutes=minutes)
            if 2020 <= dt_utc.year <= 2030:
                return dt_utc
        except:
            pass
        return None
    
    def format_datetime_for_user(self, dt_utc, include_timezone=True):
        """將UTC時間轉換為使用者時區並格式化"""
        if not dt_utc:
            return ""
        
        dt_user = dt_utc.astimezone(self.user_timezone)
        
        if include_timezone:
            tz_name = self.get_timezone_name()
            return f"{dt_user.strftime('%Y-%m-%d %H:%M:%S')} {tz_name}"
        else:
            return dt_user.strftime('%Y-%m-%d %H:%M:%S')
    
    def get_timezone_name(self):
        """取得時區名稱"""
        offset = self.user_timezone.utcoffset(datetime.now())
        if offset:
            total_seconds = int(offset.total_seconds())
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            if minutes == 0:
                return f"UTC{'+' if hours >= 0 else ''}{hours}"
            else:
                return f"UTC{'+' if hours >= 0 else ''}{hours}:{minutes:02d}"
        return "UTC"
    
    def clean_text(self, text):
        """清理文字，移除控制字符，特別處理中文字符"""
        if not text:
            return None
        
        # 移除UTF-16解碼產生的替換字符
        text = text.replace('\ufffd', '')
        
        # 移除控制字符，但保留中文字符
        # 保留基本拉丁字符、中文字符、標點符號等
        cleaned_chars = []
        for char in text:
            char_code = ord(char)
            # 保留可見字符和中文字符
            if (char_code >= 32 and char_code <= 126) or \
               (char_code >= 0x4e00 and char_code <= 0x9fff) or \
               (char_code >= 0x3400 and char_code <= 0x4dbf) or \
               char in '，。！？；：「」『』（）【】《》〈〉' or \
               char in ' \t':
                cleaned_chars.append(char)
        
        text = ''.join(cleaned_chars)
        
        # 移除多餘空白
        text = re.sub(r'\s+', ' ', text)
        return text.strip()
    
    def clean_csv_text(self, text):
        """清理CSV文字，移除換行符和特殊字符"""
        if not text:
            return ""
        # 移除換行符和回車符
        text = re.sub(r'[\r\n]+', ' ', text)
        # 移除多餘空白
        text = re.sub(r'\s+', ' ', text)
        # 移除控制字符
        text = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', text)
        return text.strip()
    
    def extract_subject_smart(self, raw_strings, html_content=None, body_content=None, file_path=None):
        """智能提取主題（基於二進制協議的通用方法）"""
        
        # 方法1: 直接從二進制文件協議中提取（最準確的方法）
        if file_path:
            binary_subject, binary_location = self.extract_subject_and_location_from_binary_protocol(file_path)
            if binary_subject:
                return binary_subject
        
        # 方法2: 從HTML內容中提取標準化主題（僅用於Google Calendar等標準格式）
        if html_content:
            html_title_patterns = [
                r'<span[^>]*itemprop="name"[^>]*>([^<]+)</span>',
                r'<title>([^<]+)</title>',
            ]
            
            for pattern in html_title_patterns:
                matches = re.findall(pattern, html_content, re.IGNORECASE)
                for match in matches:
                    cleaned_title = self.clean_text(match)
                    if cleaned_title and len(cleaned_title) >= 3:
                        return cleaned_title
        
        # 方法3: 基本文本清理（作為最後回退）
        for text in raw_strings:
            if len(text) >= 3:
                cleaned_text = self.clean_text(text)
                if cleaned_text and len(cleaned_text) >= 3:
                    return cleaned_text
        
        return None
    
    def extract_subject_and_location_from_binary_protocol(self, file_path):
        """基於.olk15Event二進制協議提取Subject和Location（純協議方法）"""
        try:
            with open(file_path, 'rb') as f:
                data = f.read()
            
            # 使用標記字節方法查找長度字段
            subject_length, location_length = self.find_field_lengths(data)
            
            if subject_length is None or location_length is None:
                print("未找到長度字段，無法解析")
                return None, None
            
            # 方法1: 查找</html>標籤的UTF-16編碼
            html_end_pattern = b'\x3c\x00\x2f\x00\x68\x00\x74\x00\x6d\x00\x6c\x00\x3e\x00'
            html_end_pos = data.find(html_end_pattern)
            
            if html_end_pos != -1:
                # 標準方法：</html>標籤後 + 回車符(0d 00)
                subject_start = html_end_pos + len(html_end_pattern) + 2
                print(f"使用</html>標籤方法，Subject開始位置: 0x{subject_start:x}")
            else:
                # 方法2: 查找 == 分隔符模式（用於某些特殊事件）
                eq_pattern = b'\x3d\x3d'  # ==
                eq_pos = data.find(eq_pattern)
                
                if eq_pos != -1:
                    # 在==之後查找UTF-16字符的開始
                    subject_start = eq_pos + 2
                    print(f"使用==分隔符方法，Subject開始位置: 0x{subject_start:x}")
                else:
                    print("未找到HTML標籤或分隔符，無法確定起始位置")
                    return None, None
            
            if subject_start >= len(data):
                return None, None
            
            # 使用長度字段精確提取Subject
            subject = None
            if subject_length > 0 and subject_start + subject_length <= len(data):
                subject_bytes = data[subject_start:subject_start + subject_length]
                subject = self.decode_utf16_bytes(subject_bytes)
            
            # 使用長度字段精確提取Location
            location = None
            if location_length > 0:
                location_start = subject_start + subject_length
                if location_start + location_length <= len(data):
                    location_bytes = data[location_start:location_start + location_length]
                    location = self.decode_utf16_bytes(location_bytes)
            else:
                location = ""  # 長度為0表示空Location
            
            return subject, location
            
        except Exception as e:
            print(f"二進制協議解析失敗: {e}")
            return None, None
    
    def find_field_lengths(self, data):
        """基於標記字節搜索Subject和Location的長度字段"""
        try:
            # 搜索Subject長度字段的標記字節: 02 00 00 1f
            subject_marker = b'\x02\x00\x00\x1f'
            
            # 在文件頭部搜索標記字節
            for pos in range(0x100, min(0x300, len(data) - 16)):
                if data[pos:pos+4] == subject_marker:
                    # 檢查後面是否有對應的標記字節 04 00 00 1f
                    if pos + 12 < len(data) and data[pos+8:pos+12] == b'\x04\x00\x00\x1f':
                        # 讀取Subject和Location長度
                        subject_len_pos = pos + 4
                        location_len_pos = pos + 12
                        
                        if subject_len_pos + 4 <= len(data) and location_len_pos + 4 <= len(data):
                            subject_len = int.from_bytes(data[subject_len_pos:subject_len_pos+4], 'little')
                            location_len = int.from_bytes(data[location_len_pos:location_len_pos+4], 'little')
                            
                            # 驗證長度是否合理（允許Location為空）
                            if (2 <= subject_len <= 500 and 0 <= location_len <= 500 and
                                self.validate_field_lengths(data, subject_len, location_len)):
                                print(f"找到標記字節長度字段 - Subject: {subject_len}字節, Location: {location_len}字節 (標記位置: 0x{pos:x})")
                                return subject_len, location_len
            
            print("未找到標記字節模式")
            return None, None
            
        except Exception as e:
            print(f"查找標記字節失敗: {e}")
            return None, None
    
    def validate_field_lengths(self, data, subject_len, location_len):
        """驗證長度字段是否對應有效的UTF-16文本"""
        try:
            # 方法1: 查找</html>標籤位置
            html_end_pattern = b'\x3c\x00\x2f\x00\x68\x00\x74\x00\x6d\x00\x6c\x00\x3e\x00'
            html_end_pos = data.find(html_end_pattern)
            
            if html_end_pos != -1:
                # 標準方法
                subject_start = html_end_pos + len(html_end_pattern) + 2
            else:
                # 方法2: 查找 == 分隔符模式
                eq_pattern = b'\x3d\x3d'  # ==
                eq_pos = data.find(eq_pattern)
                if eq_pos != -1:
                    subject_start = eq_pos + 2
                else:
                    # 如果都找不到，跳過驗證（相信長度字段）
                    return True
            
            # 檢查Subject位置是否有效
            if subject_start + subject_len > len(data):
                return False
            
            # 嘗試解碼Subject（放寬要求）
            if subject_len > 0:
                subject_bytes = data[subject_start:subject_start + subject_len]
                subject_text = self.decode_utf16_bytes(subject_bytes)
                
                if not subject_text or len(subject_text.strip()) < 1:
                    return False
            
            # 檢查Location（如果有的話）
            if location_len > 0:
                location_start = subject_start + subject_len
                if location_start + location_len > len(data):
                    return False
                
                location_bytes = data[location_start:location_start + location_len]
                location_text = self.decode_utf16_bytes(location_bytes)
                # Location可以為空，所以不檢查內容
            
            return True
            
        except Exception as e:
            print(f"驗證長度字段失敗: {e}")
            return False
            
            # 檢查Location位置是否有效
            location_start = subject_start + subject_len
            if location_start + location_len > len(data):
                return False
            
            # 嘗試解碼Location
            location_bytes = data[location_start:location_start + location_len]
            location_text = self.decode_utf16_bytes(location_bytes)
            
            # Location可以為空，但如果不為空應該是有效文本
            if location_len > 0 and location_text and len(location_text) < 2:
                return False
            
            return True
            
        except:
            return False
    
    def decode_utf16_bytes(self, byte_array):
        """解碼UTF-16字節數組（改進的邊界處理）"""
        if len(byte_array) < 2:
            return None
        
        try:
            # 首先嘗試直接解碼
            decoded = byte_array.decode('utf-16le', errors='ignore')
            
            # 清理解碼結果 - 移除控制字符和無效字符
            cleaned_chars = []
            for char in decoded:
                # 保留可打印字符和中文字符
                if (char.isprintable() or 
                    '\u4e00' <= char <= '\u9fff' or  # 中文
                    '\u3400' <= char <= '\u4dbf'):   # 中文擴展
                    cleaned_chars.append(char)
                elif char in ['\r', '\n', '\t']:
                    # 保留基本的空白字符
                    cleaned_chars.append(' ')
                else:
                    # 遇到控制字符或無效字符，可能是字段邊界
                    break
            
            result = ''.join(cleaned_chars).strip()
            
            # 進一步清理：移除末尾的重複字符模式
            result = self.clean_trailing_garbage(result)
            
            return result if len(result) >= 1 else None
            
        except:
            pass
        
        return None
    
    def clean_trailing_garbage(self, text):
        """清理末尾的垃圾字符"""
        if not text:
            return text
        
        # 移除末尾的重複特殊字符模式
        # 如 ȀȀȀ̀̀̀̀̀̀̀̀̀̀
        import re
        
        # 模式1: 移除末尾的重複Unicode控制字符
        text = re.sub(r'[\u0100-\u017f\u0300-\u036f]{3,}$', '', text)
        
        # 模式2: 移除末尾的重複特殊字符
        text = re.sub(r'[^\w\s\u4e00-\u9fff\[\]()（）【】""''.,!?;:：；，。！？-]{3,}$', '', text)
        
        # 模式3: 移除末尾的Lin, Len模式
        text = re.sub(r'Lin,\s*Len\s*$', '', text)
        
        return text.strip()
    
    def is_meaningful_subject(self, text):
        """判斷文本是否是有意義的主題"""
        if not text or len(text) < 3:
            return False
        
        # 包含中文字符
        has_chinese = any('\u4e00' <= char <= '\u9fff' for char in text)
        
        # 包含英文單詞
        has_english_words = bool(re.search(r'[A-Za-z]{2,}', text))
        
        # 包含數字和字母的合理組合
        has_reasonable_content = bool(re.search(r'[A-Za-z\u4e00-\u9fff]', text))
        
        # 不包含過多的特殊字符
        special_char_ratio = sum(1 for char in text if not (char.isalnum() or char.isspace() or char in '[]()【】-_:：')) / len(text)
        
        return (has_chinese or has_english_words or has_reasonable_content) and special_char_ratio < 0.5
    
    
    def extract_location_clean(self, raw_strings, html_content, file_path=None):
        """提取乾淨的地點（基於二進制協議的通用方法）"""
        
        # 方法1: 直接從二進制文件協議中提取（最準確的方法）
        if file_path:
            try:
                binary_subject, binary_location = self.extract_subject_and_location_from_binary_protocol(file_path)
                if binary_location:
                    return self.clean_text(binary_location)
            except:
                pass
        
        # 方法2: 從HTML內容中提取地點信息（僅用於Google Calendar等標準格式）
        if html_content:
            location_patterns = [
                r'<span[^>]*itemprop="name"[^>]*>([^<]+)</span>',
            ]
            
            for pattern in location_patterns:
                match = re.search(pattern, html_content)
                if match:
                    location = match.group(1).strip()
                    if location and len(location) >= 2:
                        return self.clean_text(location)
        
        # 方法3: 基本文本清理（作為最後回退）
        for text in raw_strings:
            if len(text) >= 2:
                cleaned_text = self.clean_text(text)
                if cleaned_text and len(cleaned_text) >= 2:
                    return cleaned_text
        
        return None
    
    def is_likely_subject(self, text):
        """判斷文本是否更像是主題而不是地點（通用方法）"""
        if not text:
            return False
        
        # 如果文本很長，更可能是主題
        if len(text) > 100:
            return True
        
        # 如果包含HTML標籤，更可能是主題
        if '<' in text and '>' in text:
            return True
        
        return False
    
    def extract_body_clean(self, html_content):
        """提取乾淨的Body內容"""
        if not html_content:
            return None
        
        # 移除HTML標籤
        text = re.sub(r'<[^>]+>', '', html_content)
        text = unescape(text)
        
        # 清理格式
        text = re.sub(r'\r\n', '\n', text)
        text = re.sub(r'\r', '\n', text)
        text = re.sub(r'\n\s*\n', '\n\n', text)
        text = re.sub(r'[ \t]+', ' ', text)
        text = text.strip()
        
        # 移除過短或無意義的內容
        if not text or len(text) < 10:
            return None
        
        # 移除常見的無用前綴
        prefixes_to_remove = [
            r'^/\*.*?\*/',
            r'^BM_BEGIN.*?BM_END',
        ]
        
        for prefix in prefixes_to_remove:
            text = re.sub(prefix, '', text, flags=re.DOTALL)
        
        text = text.strip()
        return text if text and len(text) > 10 else None
    
    def get_calendar_events_from_db(self, days=14):
        """從SQLite資料庫讀取接下來指定天數的行事曆事件，包含UID和ModDate"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            now_utc = datetime.now(timezone.utc)
            today_utc = now_utc.replace(hour=0, minute=0, second=0, microsecond=0)
            future_date_utc = today_utc + timedelta(days=days)
            
            filetime_epoch = datetime(1601, 1, 1, tzinfo=timezone.utc)
            today_minutes = int((today_utc - filetime_epoch).total_seconds() / 60)
            future_minutes = int((future_date_utc - filetime_epoch).total_seconds() / 60)
            
            print(f"查詢時間範圍 (UTC): {today_utc.strftime('%Y-%m-%d')} 到 {future_date_utc.strftime('%Y-%m-%d')}")
            print(f"使用者時區: {self.get_timezone_name()}")
            print(f"匯出天數: {days} 天")
            
            # 修改查詢以包含Calendar_UID和Record_ModDate
            query = """
            SELECT Calendar_StartDateUTC, Calendar_EndDateUTC, PathToDataFile, 
                   Calendar_UID, Record_ModDate
            FROM CalendarEvents
            WHERE Calendar_StartDateUTC >= ? AND Calendar_StartDateUTC <= ?
            ORDER BY Calendar_StartDateUTC
            """
            
            cursor.execute(query, (today_minutes, future_minutes))
            events = cursor.fetchall()
            
            print(f"找到 {len(events)} 個事件")
            conn.close()
            return events
            
        except Exception as e:
            print(f"讀取資料庫錯誤: {e}")
            return []
    
    def parse_event_file(self, file_path):
        """解析單個事件檔案"""
        try:
            with open(file_path, 'rb') as f:
                data = f.read()
        except Exception as e:
            print(f"無法讀取檔案 {file_path}: {e}")
            return None
        
        event_data = {
            'subject': None,
            'location': None,
            'organizer': None,
            'body': None,
            'start_time_utc': None,
            'end_time_utc': None,
            'duration': None
        }
        
        # 提取組織者電子郵件
        email_pattern = rb'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
        matches = re.findall(email_pattern, data)
        for match in matches:
            try:
                email = match.decode('utf-8')
                if '@' in email and not email.startswith('no-reply'):
                    event_data['organizer'] = email
                    break
            except:
                continue
        
        # 提取HTML內容
        html_content = None
        html_start = data.find(b'<\x00h\x00t\x00m\x00l\x00')
        if html_start != -1:
            html_end_pattern = b'<\x00/\x00h\x00t\x00m\x00l\x00>\x00'
            html_end = data.find(html_end_pattern, html_start)
            
            if html_end != -1:
                html_end += len(html_end_pattern)
                html_bytes = data[html_start:html_end]
                try:
                    html_content = html_bytes.decode('utf-16le', errors='ignore')
                    if '</html>' in html_content:
                        end_pos = html_content.find('</html>') + 7
                        html_content = html_content[:end_pos]
                except:
                    pass
        
        # 提取Body
        event_data['body'] = self.extract_body_clean(html_content)
        
        # 提取UTF-16字串（改進版，更好地處理中文）
        raw_strings = []
        
        # 方法1: 搜尋HTML結束後的UTF-16字串
        if html_start != -1:
            html_end_pattern = b'<\x00/\x00h\x00t\x00m\x00l\x00>\x00'
            html_end = data.find(html_end_pattern, html_start)
            if html_end != -1:
                search_start = html_end + len(html_end_pattern)
                
                # 跳過空字節
                while (search_start < len(data) and 
                       (data[search_start] == 0 or data[search_start] in [0x0d, 0x0a])):
                    search_start += 1
                
                # 收集UTF-16字串
                pos = search_start
                while pos < len(data) - 4 and len(raw_strings) < 15:
                    if (pos < len(data) - 3 and
                        data[pos] != 0 and data[pos+1] == 0 and 
                        data[pos+2] != 0 and data[pos+3] == 0):
                        
                        start = pos
                        while (pos < len(data) - 1 and 
                               data[pos] != 0 and data[pos+1] == 0):
                            pos += 2
                        
                        if pos - start >= 6:
                            try:
                                utf16_text = data[start:pos].decode('utf-16le', errors='ignore')
                                cleaned_text = self.clean_text(utf16_text)
                                
                                if cleaned_text and len(cleaned_text) >= 3:
                                    raw_strings.append(cleaned_text)
                            except:
                                pass
                        
                        while pos < len(data) and data[pos] == 0:
                            pos += 1
                    else:
                        pos += 1
        
        # 方法2: 改進的UTF-16字串搜尋（修正中文字符解碼）
        pos = 0
        while pos < len(data) - 10 and len(raw_strings) < 20:
            # 尋找UTF-16 LE模式，但更仔細地處理字節對齊
            if (pos < len(data) - 9 and
                data[pos] != 0 and data[pos+1] == 0 and
                data[pos+2] != 0 and data[pos+3] == 0):
                
                start = pos
                # 更精確地找到字串結尾
                current_pos = pos
                valid_string = True
                
                # 檢查是否為有效的UTF-16序列
                while current_pos < len(data) - 1:
                    if data[current_pos] == 0 and data[current_pos+1] == 0:
                        # 找到字串結尾
                        break
                    elif data[current_pos+1] != 0:
                        # 不是有效的UTF-16 LE格式
                        valid_string = False
                        break
                    current_pos += 2
                
                if valid_string and current_pos > start:
                    try:
                        # 使用更嚴格的UTF-16解碼
                        utf16_bytes = data[start:current_pos]
                        
                        # 確保字節數是偶數
                        if len(utf16_bytes) % 2 == 0:
                            utf16_text = utf16_bytes.decode('utf-16le', errors='replace')
                            
                            # 移除替換字符和控制字符
                            utf16_text = utf16_text.replace('\ufffd', '')
                            utf16_text = ''.join(char for char in utf16_text if ord(char) >= 32 or char in '\n\r\t')
                            
                            cleaned_text = self.clean_text(utf16_text)
                            
                            # 檢查是否包含有意義的內容
                            if (cleaned_text and len(cleaned_text) >= 3 and
                                not cleaned_text.startswith('http') and
                                cleaned_text not in raw_strings):
                                raw_strings.append(cleaned_text)
                    except Exception as e:
                        pass
                
                pos = current_pos + 2 if current_pos > start else pos + 2
            else:
                pos += 1
        
        # 提取主題和地點 - 優先使用二進制協議方法
        binary_subject, binary_location = self.extract_subject_and_location_from_binary_protocol(file_path)
        
        if binary_subject is not None:
            event_data['subject'] = binary_subject
            print(f"使用二進制協議提取Subject: {binary_subject}")
        else:
            event_data['subject'] = self.extract_subject_smart(raw_strings, html_content, event_data.get('body'), file_path)
        
        if binary_location is not None:
            event_data['location'] = binary_location
            print(f"使用二進制協議提取Location: {binary_location}")
        else:
            event_data['location'] = self.extract_location_clean(raw_strings, html_content, file_path)
        
        # 提取時間資訊
        datetime_candidates = []
        for i in range(0, len(data) - 4, 4):
            val32 = struct.unpack('<I', data[i:i+4])[0]
            if 220000000 <= val32 <= 230000000:
                dt_utc = self.minutes_since_1601_to_datetime(val32)
                if dt_utc:
                    datetime_candidates.append((val32, dt_utc))
        
        # 移除重複並按時間排序
        seen = set()
        unique_candidates = []
        for val, dt in datetime_candidates:
            if val not in seen:
                seen.add(val)
                unique_candidates.append((val, dt))
        
        unique_candidates.sort(key=lambda x: x[1])
        
        if unique_candidates:
            if len(unique_candidates) >= 2:
                # 尋找合理的時間對
                for i in range(len(unique_candidates)):
                    for j in range(i + 1, len(unique_candidates)):
                        dt1 = unique_candidates[i][1]
                        dt2 = unique_candidates[j][1]
                        
                        duration_seconds = (dt2 - dt1).total_seconds()
                        if 900 <= duration_seconds <= 28800:  # 15分鐘到8小時
                            event_data['start_time_utc'] = dt1
                            event_data['end_time_utc'] = dt2
                            event_data['duration'] = duration_seconds / 3600
                            break
                    if event_data['start_time_utc']:
                        break
                
                if not event_data['start_time_utc']:
                    event_data['start_time_utc'] = unique_candidates[0][1]
                    event_data['end_time_utc'] = unique_candidates[1][1]
                    event_data['duration'] = (event_data['end_time_utc'] - event_data['start_time_utc']).total_seconds() / 3600
            else:
                event_data['start_time_utc'] = unique_candidates[0][1]
        
        return event_data
    
    def process_events(self, days=14):
        """處理所有事件"""
        db_events = self.get_calendar_events_from_db(days)
        
        if not db_events:
            print("沒有找到事件")
            return []
        
        processed_events = []
        
        for start_minutes, end_minutes, path_to_data_file, calendar_uid, record_mod_date in db_events:
            print(f"\n處理事件: {path_to_data_file}")
            
            full_path = os.path.join(self.outlook_data_path, path_to_data_file)
            
            if not os.path.exists(full_path):
                print(f"檔案不存在: {full_path}")
                continue
            
            event_data = self.parse_event_file(full_path)
            
            if event_data:
                # 添加資料庫欄位
                event_data['calendar_uid'] = calendar_uid
                event_data['record_mod_date'] = record_mod_date
                event_data['path_to_data_file'] = path_to_data_file
                
                # 始終使用資料庫中的UTC時間（最可靠）
                event_data['start_time_utc'] = self.minutes_since_1601_to_datetime(start_minutes)
                event_data['end_time_utc'] = self.minutes_since_1601_to_datetime(end_minutes)
                
                # 計算持續時間
                if event_data['start_time_utc'] and event_data['end_time_utc'] and not event_data['duration']:
                    event_data['duration'] = (event_data['end_time_utc'] - event_data['start_time_utc']).total_seconds() / 3600
                
                processed_events.append(event_data)
                
                # 顯示解析結果
                print(f"  Subject: {event_data['subject'] or '(Unknown)'}")
                print(f"  Location: {event_data['location'] or '(Unknown)'}")
                print(f"  Organizer: {event_data['organizer'] or '(Unknown)'}")
        
        return processed_events
    
    def export_to_csv(self, events, output_file="data/dump_outlook_calendar.csv"):
        """將事件匯出為CSV檔案（包含Calendar_UID和Record_ModDate，修正格式問題）"""
        if not events:
            print("沒有事件可匯出")
            return
        
        with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
            # 包含Calendar_UID和Record_ModDate欄位
            fieldnames = [
                'Calendar_UID', 'Record_ModDate', 'Subject', 'Location', 'Organizer', 
                'Duration', 'Starts', 'Ends', 'Starts_UTC', 'Ends_UTC', 'Body', 'PathToDataFile'
            ]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            writer.writeheader()
            
            for event in events:
                # 格式化時間
                starts_user = self.format_datetime_for_user(event['start_time_utc'], include_timezone=False)
                ends_user = self.format_datetime_for_user(event['end_time_utc'], include_timezone=False)
                starts_utc = event['start_time_utc'].strftime('%Y-%m-%d %H:%M:%S UTC') if event['start_time_utc'] else ""
                ends_utc = event['end_time_utc'].strftime('%Y-%m-%d %H:%M:%S UTC') if event['end_time_utc'] else ""
                
                # 清理所有文字欄位以避免CSV格式問題
                writer.writerow({
                    'Calendar_UID': self.clean_csv_text(event['calendar_uid'] or ''),
                    'Record_ModDate': event['record_mod_date'] or '',
                    'Subject': self.clean_csv_text(event['subject'] or ''),
                    'Location': self.clean_csv_text(event['location'] or ''),
                    'Organizer': self.clean_csv_text(event['organizer'] or ''),
                    'Duration': f"{event['duration']:.1f}" if event['duration'] else '',
                    'Starts': starts_user,
                    'Ends': ends_user,
                    'Starts_UTC': starts_utc,
                    'Ends_UTC': ends_utc,
                    'Body': self.clean_csv_text(event['body'] or ''),
                    'PathToDataFile': self.clean_csv_text(event['path_to_data_file'] or '')
                })
        
        print(f"\n已匯出 {len(events)} 個事件到 {output_file}")
        print(f"時區設定: {self.get_timezone_name()}")
        print("CSV欄位包含: Calendar_UID, Record_ModDate, Subject, Location, Organizer, Duration, Starts, Ends, Starts_UTC, Ends_UTC, Body, PathToDataFile")

def main():
    parser = argparse.ArgumentParser(description='修正版完整時區感知Mac Outlook Calendar Reader')
    parser.add_argument('--timezone', '-tz', default='UTC+8', 
                       help='使用者時區 (例如: UTC+8, UTC-5, UTC+0)')
    parser.add_argument('--days', '-d', type=int, default=14,
                       help='匯出天數 (預設: 14天)')
    
    args = parser.parse_args()
    
    print("修正版完整時區感知Mac Outlook Calendar Reader")
    print("包含Calendar_UID和Record_ModDate欄位")
    print("=" * 60)
    
    reader = CompleteFixedTimeZoneOutlookParser(user_timezone=args.timezone)
    
    if not os.path.exists(reader.db_path):
        print(f"錯誤: 找不到Outlook資料庫: {reader.db_path}")
        sys.exit(1)
    
    events = reader.process_events(args.days)
    
    if events:
        reader.export_to_csv(events)
    else:
        print("沒有找到任何事件")

if __name__ == "__main__":
    main()
