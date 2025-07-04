#!/usr/bin/env python3
"""
改進版 Outlook Calendar to Google Calendar 同步器
支援事件更新、去重複、錯誤處理、強制更新
"""

import pandas as pd
import datetime
import json
import os
import sys
import re
import argparse
from pathlib import Path
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

class OutlookToGoogleCalendarSync:
    def __init__(self, csv_path="data/dump_outlook_calendar.csv", 
                 client_secret_file="data/client_secret.json",
                 calendar_id="OutlookMacSync",
                 force_update=False,
                 mark_deleted=True,
                 cleanup_days=2,
                 enable_cleanup=True):
        self.csv_path = csv_path
        self.cache_path = "data/sync_cache.json"
        self.token_path = "data/token.json"
        self.client_secret_file = client_secret_file
        self.calendar_id = calendar_id
        self.scopes = ['https://www.googleapis.com/auth/calendar']
        self.service = None
        self.cache = {}
        self.force_update = force_update
        self.mark_deleted = mark_deleted
        self.cleanup_days = cleanup_days
        self.enable_cleanup = enable_cleanup
        
    def authenticate(self):
        """Google Calendar API 認證"""
        creds = None
        
        # 檢查是否有已存在的 token
        if os.path.exists(self.token_path):
            try:
                creds = Credentials.from_authorized_user_file(self.token_path, self.scopes)
                print(f"📁 載入現有憑證: {self.token_path}")
                
                # 顯示憑證狀態
                if creds.expired:
                    if creds.refresh_token:
                        print("⏰ 憑證已過期，但有刷新令牌")
                    else:
                        print("❌ 憑證已過期且無刷新令牌")
                else:
                    # 計算憑證剩餘時間，並提前刷新
                    if hasattr(creds, 'expiry') and creds.expiry:
                        from datetime import datetime, timezone
                        now = datetime.now(timezone.utc)
                        
                        # 確保 expiry 也是 timezone-aware
                        expiry = creds.expiry
                        if expiry.tzinfo is None:
                            # 如果 expiry 是 naive，假設它是 UTC 時間
                            expiry = expiry.replace(tzinfo=timezone.utc)
                        
                        remaining = expiry - now
                        
                        if remaining.total_seconds() > 0:
                            # 如果剩餘時間少於10分鐘，提前刷新
                            if remaining.total_seconds() < 600:  # 10分鐘
                                print("⚠️ Access Token 即將過期，提前刷新...")
                                if creds.refresh_token:
                                    try:
                                        creds.refresh(Request())
                                        print("✅ 提前刷新成功")
                                        # 立即保存刷新後的憑證
                                        with open(self.token_path, 'w') as token:
                                            token.write(creds.to_json())
                                        print("💾 已保存刷新後的憑證")
                                    except Exception as e:
                                        print(f"⚠️ 提前刷新失敗: {e}")
                            else:
                                hours = int(remaining.total_seconds() // 3600)
                                minutes = int((remaining.total_seconds() % 3600) // 60)
                                print(f"✅ Access Token 有效，剩餘: {hours}h{minutes}m")
                        else:
                            print("⏰ Access Token 已過期")
                    else:
                        print("✅ Access Token 有效")
                        
            except Exception as e:
                print(f"❌ 載入憑證失敗: {e}")
                # 刪除無效的 token 檔案
                if os.path.exists(self.token_path):
                    os.remove(self.token_path)
                    print("🗑️ 已刪除無效的憑證檔案")
                creds = None
        
        # 如果沒有有效的憑證，進行 OAuth 流程
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    print("🔄 正在刷新 Access Token...")
                    creds.refresh(Request())
                    print("✅ Access Token 刷新成功")
                    
                    # 立即保存刷新後的憑證
                    with open(self.token_path, 'w') as token:
                        token.write(creds.to_json())
                    print("💾 已保存刷新後的憑證")
                    
                except Exception as e:
                    print(f"❌ 刷新 Access Token 失敗: {e}")
                    print("💡 可能的原因:")
                    print("   • Refresh Token 已過期（超過6個月未使用）")
                    print("   • 用戶撤銷了應用授權")
                    print("   • Google 帳戶密碼已更改")
                    print("   • 網路連線問題")
                    creds = None
            
            if not creds:
                if not os.path.exists(self.client_secret_file):
                    print(f"❌ 錯誤: 找不到 Google API 憑證檔案: {self.client_secret_file}")
                    print("📋 請從 Google Cloud Console 下載 OAuth 2.0 憑證檔案")
                    print("🔗 https://console.cloud.google.com/apis/credentials")
                    sys.exit(1)
                
                print("🔐 開始 OAuth 2.0 授權流程...")
                print("💡 提示：授權後憑證將保存到 token.json")
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.client_secret_file, self.scopes)
                creds = flow.run_local_server(port=0)
            
            # 儲存憑證供下次使用
            with open(self.token_path, 'w') as token:
                token.write(creds.to_json())
            print(f"💾 憑證已保存到: {self.token_path}")
        
        self.service = build('calendar', 'v3', credentials=creds)
        print("✅ Google Calendar API 認證成功")
        
        # 顯示憑證維護提示
        self._show_maintenance_tips(creds)
    
    def _show_maintenance_tips(self, creds):
        """顯示憑證維護提示"""
        print("\n💡 憑證維護資訊:")
        
        if creds and creds.refresh_token:
            print("   ✅ Refresh Token: 存在")
            print("   • 有效期：6個月（需定期使用保持活躍）")
            print("   • Access Token 會自動刷新（1小時有效期）")
            print("   • 建議每月至少運行一次同步")
            
            # 檢查憑證檔案的最後修改時間
            if os.path.exists(self.token_path):
                import time
                mtime = os.path.getmtime(self.token_path)
                days_ago = (time.time() - mtime) / (24 * 3600)
                
                print(f"   📅 憑證最後更新：{int(days_ago)} 天前")
                
                if days_ago > 150:  # 5個月
                    print("   🚨 警告：憑證超過5個月未更新，接近過期！")
                    print("   💡 建議：立即運行同步以刷新憑證")
                elif days_ago > 90:  # 3個月
                    print("   ⚠️ 注意：憑證超過3個月未更新")
                    print("   💡 建議：近期內運行同步")
                elif days_ago > 30:  # 1個月
                    print("   ℹ️ 憑證超過1個月未更新")
                else:
                    print("   ✅ 憑證狀態良好")
        else:
            print("   ❌ Refresh Token: 不存在")
            print("   ⚠️ 警告：無法自動刷新 Access Token")
            print("   💡 過期後需要重新完整授權")
        
        print("\n🔄 自動刷新機制:")
        print("   • Access Token 剩餘時間 < 10分鐘時自動提前刷新")
        print("   • 刷新成功後立即保存新憑證")
        print("   • 刷新失敗時會提示重新授權")
        print("   • 透明處理，用戶無感知")
    
    def cleanup_expired_events(self, days_threshold=2):
        """清理過期的事件
        
        Args:
            days_threshold (int): 過期天數閾值，預設2天
        """
        try:
            from datetime import datetime, timezone, timedelta
            
            # 計算過期時間點（前天 23:59:59）
            cutoff_date = datetime.now(timezone.utc) - timedelta(days=days_threshold)
            cutoff_str = cutoff_date.strftime('%Y-%m-%dT%H:%M:%SZ')
            
            print(f"🗑️ 開始清理 {days_threshold} 天前的過期事件...")
            print(f"📅 清理截止時間: {cutoff_date.strftime('%Y-%m-%d %H:%M:%S UTC')}")
            
            # 搜索過期事件
            events_result = self.service.events().list(
                calendarId=self.calendar_id,
                timeMax=cutoff_str,  # 結束時間在截止時間之前的事件
                maxResults=2500,
                singleEvents=True,
                orderBy='startTime'
            ).execute()
            
            expired_events = events_result.get('items', [])
            
            if not expired_events:
                print("✅ 沒有找到需要清理的過期事件")
                return
            
            print(f"🔍 找到 {len(expired_events)} 個過期事件")
            
            # 刪除過期事件
            deleted_count = 0
            failed_count = 0
            
            for event in expired_events:
                try:
                    event_id = event['id']
                    event_title = event.get('summary', '無標題')
                    event_start = event.get('start', {}).get('dateTime', event.get('start', {}).get('date', '未知時間'))
                    
                    # 檢查是否是 Outlook 同步的事件（通過描述中的標記識別）
                    description = event.get('description', '')
                    
                    # 檢查多種可能的標記格式
                    is_outlook_event = (
                        'Outlook UID:' in description or 
                        '[OutlookMacSync]' in description or
                        'Outlook Calendar UID:' in description or
                        '[Outlook Calendar UID:' in description
                    )
                    
                    if is_outlook_event:
                        self.service.events().delete(
                            calendarId=self.calendar_id,
                            eventId=event_id
                        ).execute()
                        
                        deleted_count += 1
                        print(f"🗑️ 已刪除: {event_title} ({event_start})")
                    else:
                        print(f"⏭️ 跳過非同步事件: {event_title}")
                        
                except Exception as e:
                    failed_count += 1
                    print(f"❌ 刪除失敗: {event_title} - {e}")
            
            print(f"\n🎉 過期事件清理完成!")
            print(f"✅ 成功刪除: {deleted_count} 個事件")
            if failed_count > 0:
                print(f"❌ 刪除失敗: {failed_count} 個事件")
                
        except Exception as e:
            print(f"❌ 清理過期事件時發生錯誤: {e}")
    
    def setup_outlook_calendar(self):
        """設定或創建 OutlookMacSync 日曆"""
        try:
            # 如果 calendar_id 是 "OutlookMacSync"，需要找到或創建這個日曆
            if self.calendar_id == "OutlookMacSync":
                print("🔍 搜索 OutlookMacSync 日曆...")
                
                # 列出所有日曆
                calendars_result = self.service.calendarList().list().execute()
                calendars = calendars_result.get('items', [])
                
                # 尋找 OutlookMacSync 日曆
                outlook_calendar = None
                for calendar in calendars:
                    if calendar.get('summary') == 'OutlookMacSync':
                        outlook_calendar = calendar
                        break
                
                if outlook_calendar:
                    self.calendar_id = outlook_calendar['id']
                    print(f"✅ 找到現有的 OutlookMacSync 日曆")
                    print(f"📅 日曆 ID: {self.calendar_id}")
                else:
                    # 創建新的日曆
                    print("📅 創建新的 OutlookMacSync 日曆...")
                    calendar_body = {
                        'summary': 'OutlookMacSync',
                        'description': '從 Mac Outlook 同步的行事曆事件\n\n此日曆包含從 Microsoft Outlook for Mac 自動同步的事件。\n請勿手動修改此日曆中的事件，因為它們會在下次同步時被覆蓋。',
                        'timeZone': 'Asia/Taipei'
                    }
                    
                    created_calendar = self.service.calendars().insert(body=calendar_body).execute()
                    self.calendar_id = created_calendar['id']
                    
                    print(f"✅ 成功創建 OutlookMacSync 日曆")
                    print(f"📅 日曆 ID: {self.calendar_id}")
                    
                    # 設定日曆顏色（可選）
                    try:
                        calendar_list_entry = {
                            'id': self.calendar_id,
                            'colorId': '9'  # 藍色
                        }
                        self.service.calendarList().patch(
                            calendarId=self.calendar_id, 
                            body=calendar_list_entry
                        ).execute()
                        print("🎨 設定日曆顏色為藍色")
                    except Exception as e:
                        print(f"⚠️ 設定日曆顏色失敗: {e}")
            
            else:
                print(f"📅 使用指定的日曆: {self.calendar_id}")
                
        except Exception as e:
            print(f"❌ 設定日曆時發生錯誤: {e}")
            print("💡 將使用主要日曆作為備選")
            self.calendar_id = "primary"
    
    def load_cache(self):
        """載入本地快取"""
        if os.path.exists(self.cache_path):
            try:
                with open(self.cache_path, "r", encoding='utf-8') as f:
                    self.cache = json.load(f)
                print(f"📁 載入快取: {len(self.cache)} 個事件")
            except Exception as e:
                print(f"載入快取失敗: {e}")
                self.cache = {}
        else:
            self.cache = {}
    
    def save_cache(self):
        """儲存本地快取"""
        try:
            with open(self.cache_path, "w", encoding='utf-8') as f:
                json.dump(self.cache, f, ensure_ascii=False, indent=2)
            print(f"💾 快取已儲存: {len(self.cache)} 個事件")
        except Exception as e:
            print(f"儲存快取失敗: {e}")
    
    def detect_deleted_events(self, current_events_df):
        """檢測已刪除的事件（排除超出時間範圍的事件）"""
        if not self.cache:
            print("ℹ️ 快取為空，無法檢測刪除事件")
            return []
        
        # 從當前CSV中提取所有Calendar_UID
        current_uids = set(current_events_df['Calendar_UID'].astype(str))
        print(f"🔍 當前CSV中有 {len(current_uids)} 個事件")
        
        # 計算當前匯出的時間範圍
        if not current_events_df.empty:
            # 從CSV中的UTC時間計算範圍
            start_times = pd.to_datetime(current_events_df['Starts_UTC'])
            current_range_start = start_times.min().date()
            current_range_end = start_times.max().date()
            print(f"🔍 當前匯出範圍: {current_range_start} 到 {current_range_end}")
        else:
            print("⚠️ 當前CSV為空，無法確定時間範圍")
            return []
        
        # 從快取中找出不再存在於當前CSV的事件
        deleted_events = []
        cache_uids = set(self.cache.keys())
        print(f"🔍 快取中有 {len(cache_uids)} 個事件")
        
        for outlook_uid in cache_uids:
            if outlook_uid not in current_uids:
                # 檢查這個事件是否可能只是超出了時間範圍
                is_likely_out_of_range = self.check_if_event_out_of_range(
                    outlook_uid, current_range_start, current_range_end
                )
                
                if not is_likely_out_of_range:
                    # 只有當事件不是因為超出範圍才被認為是真正刪除
                    deleted_events.append({
                        'outlook_uid': outlook_uid,
                        'record_moddate': self.cache[outlook_uid]  # 這是timestamp，不是Google Event ID
                    })
                else:
                    print(f"⏰ 跳過超出範圍的事件: {outlook_uid[:30]}...")
        
        if deleted_events:
            print(f"🔍 檢測到真正刪除的事件: {len(deleted_events)}")
        
        return deleted_events
    
    def check_if_event_out_of_range(self, outlook_uid, current_range_start, current_range_end):
        """檢查事件是否為過去事件（過去事件不應被標記為刪除）"""
        try:
            # 獲取Google Calendar中的事件來確定其時間
            from datetime import datetime, timedelta, date
            
            # 擴大搜索範圍到前後各30天
            extended_start = current_range_start - timedelta(days=30)
            extended_end = current_range_end + timedelta(days=30)
            
            time_min = datetime.combine(extended_start, datetime.min.time()).isoformat() + 'Z'
            time_max = datetime.combine(extended_end, datetime.min.time()).isoformat() + 'Z'
            
            events_result = self.service.events().list(
                calendarId=self.calendar_id,
                timeMin=time_min,
                timeMax=time_max,
                singleEvents=True,
                maxResults=2500
            ).execute()
            
            google_events = events_result.get('items', [])
            
            # 在Google Calendar事件中搜索包含此Outlook UID的事件
            for google_event in google_events:
                description = google_event.get('description', '')
                if f"Outlook UID: {outlook_uid}" in description:
                    # 找到了對應的Google Calendar事件，檢查其時間
                    event_start = google_event.get('start', {})
                    if 'dateTime' in event_start:
                        event_date = pd.to_datetime(event_start['dateTime']).date()
                    elif 'date' in event_start:
                        event_date = pd.to_datetime(event_start['date']).date()
                    else:
                        continue
                    
                    today = date.today()
                    
                    # 如果是過去的事件，認為是超出範圍（不應刪除）
                    if event_date < today:
                        print(f"📅 事件 {outlook_uid[:20]}... 在 {event_date}（過去），跳過刪除檢測")
                        return True
                    else:
                        # 未來事件但不在CSV中，可能是真正被刪除
                        print(f"🔮 事件 {outlook_uid[:20]}... 在 {event_date}（未來），檢查是否被刪除")
                        return False
            
            # 如果在Google Calendar中找不到事件，保守處理
            print(f"❓ 事件 {outlook_uid[:20]}... 在Google Calendar中找不到，跳過刪除檢測")
            return True  # 保守處理：不標記為刪除
            
        except Exception as e:
            print(f"⚠️ 檢查事件時間時發生錯誤: {e}")
            # 發生錯誤時，保守處理：不標記為刪除
            return True
    
    def mark_deleted_events(self, deleted_events):
        """標記已刪除的事件（通過搜索Google Calendar找到對應事件）"""
        marked_count = 0
        cleaned_count = 0
        
        if not deleted_events:
            return 0
        
        try:
            # 獲取Google Calendar中的所有事件（接下來7天）
            from datetime import datetime, timedelta
            now = datetime.utcnow()
            time_min = (now - timedelta(days=1)).isoformat() + 'Z'  # 包含昨天的事件
            time_max = (now + timedelta(days=7)).isoformat() + 'Z'
            
            events_result = self.service.events().list(
                calendarId=self.calendar_id,
                timeMin=time_min,
                timeMax=time_max,
                singleEvents=True,
                maxResults=2500  # 增加搜索範圍
            ).execute()
            
            google_events = events_result.get('items', [])
            print(f"🔍 在Google Calendar中找到 {len(google_events)} 個事件")
            
            # 為每個已刪除的Outlook事件尋找對應的Google Calendar事件
            for deleted_event in deleted_events:
                outlook_uid = deleted_event['outlook_uid']
                found_event = None
                
                # 在Google Calendar事件中搜索包含此Outlook UID的事件
                for google_event in google_events:
                    if 'description' in google_event and outlook_uid in google_event['description']:
                        found_event = google_event
                        break
                
                if found_event:
                    current_title = found_event.get('summary', 'Untitled Event')
                    
                    # 如果標題還沒有被標記為已刪除
                    if not current_title.startswith('[DELETED]'):
                        # 更新標題
                        new_title = f"[DELETED] {current_title}"
                        found_event['summary'] = new_title
                        
                        # 更新事件描述，添加刪除信息
                        current_description = found_event.get('description', '')
                        deletion_note = f"\\n\\n⚠️ 此事件已從Outlook中刪除 (刪除時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')})"
                        found_event['description'] = current_description + deletion_note
                        
                        # 更新Google Calendar事件
                        updated_event = self.service.events().update(
                            calendarId=self.calendar_id,
                            eventId=found_event['id'],
                            body=found_event
                        ).execute()
                        
                        print(f"🗑️ 標記已刪除事件: {current_title}")
                        marked_count += 1
                    else:
                        print(f"ℹ️ 事件已標記為刪除: {current_title}")
                else:
                    print(f"🧹 未找到對應的Google Calendar事件: {outlook_uid[:30]}...")
                    cleaned_count += 1
                
                # 從快取中移除已刪除的事件
                if outlook_uid in self.cache:
                    del self.cache[outlook_uid]
            
            if cleaned_count > 0:
                print(f"🧹 已清理 {cleaned_count} 個無法找到的事件")
            
        except Exception as e:
            print(f"❌ 標記刪除事件時發生錯誤: {e}")
        
        return marked_count
    
    def parse_datetime(self, datetime_str):
        """解析時間字串為 RFC3339 格式"""
        if not datetime_str or pd.isna(datetime_str):
            return None
        
        try:
            # 解析時間字串
            dt = pd.to_datetime(datetime_str)
            
            # 轉換為 RFC3339 格式
            if dt.tzinfo is None:
                # 如果沒有時區信息，假設是 UTC
                dt = dt.replace(tzinfo=datetime.timezone.utc)
            
            return dt.isoformat()
        except Exception as e:
            print(f"時間解析錯誤: {datetime_str} - {e}")
            return None
    
    def create_event_body(self, row):
        """創建 Google Calendar 事件主體"""
        # 基本事件信息
        event_body = {
            'summary': str(row['Subject']) if pd.notna(row['Subject']) else 'Untitled Event',
            'start': {
                'dateTime': self.parse_datetime(row['Starts_UTC']),
                'timeZone': 'UTC'
            },
            'end': {
                'dateTime': self.parse_datetime(row['Ends_UTC']),
                'timeZone': 'UTC'
            }
        }
        
        # 添加地點信息
        if pd.notna(row['Location']) and str(row['Location']).strip():
            event_body['location'] = str(row['Location'])
        
        # 添加描述信息
        description_parts = []
        
        # 添加同步標記（用於清理識別）
        description_parts.append("[OutlookMacSync] 此事件由 Mac Outlook 自動同步")
        
        # 添加組織者信息
        if pd.notna(row['Organizer']) and str(row['Organizer']).strip():
            description_parts.append(f"組織者: {row['Organizer']}")
        
        # 添加 Outlook UID（用於識別）
        description_parts.append(f"Outlook UID: {row['Calendar_UID']}")
        
        # 添加 Body 內容
        if pd.notna(row['Body']) and str(row['Body']).strip():
            description_parts.append("\\n內容:")
            description_parts.append(str(row['Body']))
        
        if description_parts:
            event_body['description'] = '\\n'.join(description_parts)
        
        return event_body
    
    def parse_datetime(self, datetime_str):
        """解析時間字串為 RFC3339 格式"""
        if not datetime_str or pd.isna(datetime_str):
            return None
        """解析時間字串為 RFC3339 格式"""
        if not datetime_str or pd.isna(datetime_str):
            return None
        
        try:
            # 移除 " UTC" 後綴（如果存在）
            clean_str = str(datetime_str).replace(' UTC', '').strip()
            
            # 嘗試解析不同格式
            formats = [
                '%Y-%m-%d %H:%M:%S',
                '%Y-%m-%d %H:%M',
                '%Y-%m-%dT%H:%M:%S',
                '%Y-%m-%dT%H:%M:%SZ',
            ]
            
            for fmt in formats:
                try:
                    dt = datetime.datetime.strptime(clean_str, fmt)
                    # 假設輸入是 UTC 時間
                    dt = dt.replace(tzinfo=datetime.timezone.utc)
                    return dt.isoformat()
                except ValueError:
                    continue
            
            print(f"⚠️  無法解析時間格式: {datetime_str}")
            return None
            
        except Exception as e:
            print(f"⚠️  時間解析錯誤: {e}")
            return None
    
    def clean_text(self, text):
        """清理文字內容"""
        if not text or pd.isna(text):
            return ""
        
        text = str(text).strip()
        # 移除過長的內容（Google Calendar 有限制）
        if len(text) > 8000:
            text = text[:8000] + "...(內容已截斷)"
        
        return text
    
    def generate_event_id(self, calendar_uid):
        """生成 Google Calendar 事件 ID"""
        # Google Calendar 事件 ID 要求：
        # - 只能包含小寫字母、數字和連字符
        # - 長度 5-1024 字符
        # - 不能以數字開頭
        # - 不能以連字符結尾
        
        import hashlib
        
        # 對於所有 UID，統一使用 hash 來生成穩定且符合格式的 ID
        original_uid = str(calendar_uid)
        
        # 使用 MD5 hash 生成固定長度的 ID
        hash_obj = hashlib.md5(original_uid.encode('utf-8'))
        hash_hex = hash_obj.hexdigest()
        
        # 根據原始 UID 類型添加前綴，確保不以數字開頭
        if '@google.com' in original_uid:
            clean_uid = f"google-{hash_hex}"
        elif original_uid.startswith('Meetings-'):
            clean_uid = f"amazon-{hash_hex}"
        elif len(original_uid) > 50 and original_uid.startswith('040000008200E00074C5B7101A82E008'):
            clean_uid = f"exchange-{hash_hex}"
        elif '-' in original_uid and len(original_uid) == 36:  # GUID format
            clean_uid = f"guid-{hash_hex}"
        else:
            clean_uid = f"outlook-{hash_hex}"
        
        # 確保不以數字開頭（雖然我們已經加了前綴，但雙重保險）
        if clean_uid[0].isdigit():
            clean_uid = f"event-{clean_uid}"
        
        # 確保長度不超過 64 字符（安全範圍）
        if len(clean_uid) > 64:
            clean_uid = clean_uid[:64]
        
        # 確保不以連字符結尾
        clean_uid = clean_uid.rstrip('-')
        
        return clean_uid
    
    def create_or_update_event(self, row):
        """創建或更新 Google Calendar 事件"""
        try:
            calendar_uid = str(row['Calendar_UID'])
            record_moddate = str(row['Record_ModDate'])
            subject = self.clean_text(row['Subject'])
            location = self.clean_text(row['Location'])
            organizer = self.clean_text(row['Organizer'])
            starts_utc = self.parse_datetime(row['Starts_UTC'])
            ends_utc = self.parse_datetime(row['Ends_UTC'])
            body = self.clean_text(row['Body'])
            
            # 檢查必要欄位
            if not subject:
                subject = "(無主題)"
            
            if not starts_utc or not ends_utc:
                print(f"⚠️  跳過事件 '{subject}': 時間資訊不完整")
                return False
            
            # 檢查是否需要更新
            cache_key = calendar_uid
            if not self.force_update and cache_key in self.cache and self.cache[cache_key] == record_moddate:
                print(f"⏭️  跳過 '{subject}': 未變更")
                return True
            
            if self.force_update:
                print(f"🔄 強制更新 '{subject}'")
            elif cache_key in self.cache:
                print(f"🔄 檢測到變更，更新 '{subject}'")
            else:
                print(f"➕ 新事件 '{subject}'")
            
            # 準備事件資料
            event_body = {
                'summary': subject,
                'start': {'dateTime': starts_utc, 'timeZone': 'UTC'},
                'end': {'dateTime': ends_utc, 'timeZone': 'UTC'},
                'reminders': {'useDefault': True},
                # 在描述中加入 Calendar_UID 以便識別
                'description': f"[Outlook Calendar UID: {calendar_uid}]\n\n{body}" if body else f"[Outlook Calendar UID: {calendar_uid}]"
            }
            
            # 可選欄位
            if location:
                event_body['location'] = location
            
            if organizer and '@' in organizer:
                event_body['organizer'] = {'email': organizer}
            
            # 搜尋是否已存在相同的事件（通過描述中的 UID）
            try:
                # 搜尋包含此 Calendar_UID 的事件
                events_result = self.service.events().list(
                    calendarId=self.calendar_id,
                    q=f"Outlook Calendar UID: {calendar_uid}",
                    maxResults=10
                ).execute()
                
                events = events_result.get('items', [])
                existing_event = None
                
                # 找到匹配的事件
                for event in events:
                    if 'description' in event and calendar_uid in event['description']:
                        existing_event = event
                        break
                
                if existing_event:
                    # 更新現有事件
                    updated_event = self.service.events().update(
                        calendarId=self.calendar_id,
                        eventId=existing_event['id'],
                        body=event_body
                    ).execute()
                    
                    print(f"🔄 更新事件: {subject}")
                else:
                    # 創建新事件（不指定 ID，讓 Google 自動生成）
                    created_event = self.service.events().insert(
                        calendarId=self.calendar_id,
                        body=event_body
                    ).execute()
                    
                    print(f"➕ 創建事件: {subject}")
                
            except HttpError as e:
                print(f"❌ API 錯誤: {e}")
                return False
            
            # 更新快取
            self.cache[cache_key] = record_moddate
            return True
            
        except Exception as e:
            print(f"❌ 處理事件失敗: {e}")
            return False
    
    def sync_events(self):
        """同步所有事件"""
        # 設定 OutlookMacSync 日曆
        self.setup_outlook_calendar()
        
        # 檢查 CSV 檔案
        if not os.path.exists(self.csv_path):
            print(f"❌ 找不到 CSV 檔案: {self.csv_path}")
            print("請先執行 Outlook 行事曆讀取器生成 CSV 檔案")
            return False
        
        try:
            # 讀取 CSV
            df = pd.read_csv(self.csv_path)
            print(f"📊 讀取 CSV: {len(df)} 個事件")
            
            # 檢查必要欄位
            required_columns = ['Calendar_UID', 'Record_ModDate', 'Subject', 'Starts_UTC', 'Ends_UTC']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                print(f"❌ CSV 檔案缺少必要欄位: {missing_columns}")
                return False
            
            # 檢測已刪除的事件（如果啟用）
            if self.mark_deleted:
                deleted_events = self.detect_deleted_events(df)
                if deleted_events:
                    print(f"\n🗑️ 檢測到 {len(deleted_events)} 個已刪除的事件")
                    for event in deleted_events:
                        print(f"   - {event['outlook_uid'][:30]}...")
                    marked_count = self.mark_deleted_events(deleted_events)
                    if marked_count > 0:
                        print(f"✅ 已標記 {marked_count} 個刪除事件")
                    else:
                        print("ℹ️ 所有已刪除的事件都已處理（可能已不存在於Google Calendar中）")
                else:
                    print("\n✅ 沒有檢測到已刪除的事件")
            
            # 同步事件
            success_count = 0
            error_count = 0
            
            for index, row in df.iterrows():
                print(f"\n處理事件 {index + 1}/{len(df)}")
                
                if self.create_or_update_event(row):
                    success_count += 1
                else:
                    error_count += 1
                
                # 每 10 個事件儲存一次快取
                if (index + 1) % 10 == 0:
                    self.save_cache()
            
            # 最終儲存快取
            self.save_cache()
            
            print(f"\n🎉 同步完成!")
            print(f"✅ 成功: {success_count} 個事件")
            print(f"❌ 失敗: {error_count} 個事件")
            
            # 清理過期事件
            if self.enable_cleanup and self.cleanup_days > 0:
                print(f"\n" + "="*50)
                self.cleanup_expired_events(days_threshold=self.cleanup_days)
            else:
                print(f"\nℹ️ 過期事件清理已停用")
            
            return True
            
        except Exception as e:
            print(f"❌ 同步失敗: {e}")
            return False

def main():
    # 解析命令行參數
    parser = argparse.ArgumentParser(description='Outlook Calendar to Google Calendar 同步器')
    parser.add_argument('--force', '-f', action='store_true', 
                       help='強制更新所有事件，忽略快取檢查')
    parser.add_argument('--clear-cache', action='store_true',
                       help='清除同步快取檔案')
    parser.add_argument('--mark-deleted', action='store_true', default=True,
                       help='標記已刪除的事件（預設啟用）')
    parser.add_argument('--no-mark-deleted', action='store_true',
                       help='不標記已刪除的事件')
    parser.add_argument('--days', '-d', type=int, default=14,
                       help='同步天數，應與Outlook匯出天數一致 (預設: 14天)')
    parser.add_argument('--cleanup-days', type=int, default=2,
                       help='自動清理多少天前的過期事件 (預設: 2天，設為0則停用)')
    parser.add_argument('--no-cleanup', action='store_true',
                       help='停用自動清理過期事件')
    args = parser.parse_args()
    
    print("Outlook Calendar to Google Calendar 同步器")
    print("=" * 50)
    
    if args.force:
        print("🔄 強制更新模式：將更新所有事件")
    
    print(f"📅 同步範圍：{args.days} 天")
    
    # 處理刪除標記選項
    mark_deleted = args.mark_deleted and not args.no_mark_deleted
    if mark_deleted:
        print("🗑️ 刪除檢測：已啟用（將標記已刪除的事件）")
    else:
        print("ℹ️ 刪除檢測：已停用")
    
    # 處理清理選項
    enable_cleanup = not args.no_cleanup and args.cleanup_days > 0
    if enable_cleanup:
        print(f"🧹 自動清理：已啟用（清理 {args.cleanup_days} 天前的過期事件）")
    else:
        print("ℹ️ 自動清理：已停用")
    
    if args.clear_cache:
        cache_file = "sync_cache.json"
        if os.path.exists(cache_file):
            os.remove(cache_file)
            print(f"🗑️  已清除快取檔案: {cache_file}")
        else:
            print("ℹ️  快取檔案不存在")
    
    # 檢查是否有 CSV 檔案
    csv_files = [
        "data/dump_outlook_calendar.csv",
        "dump_outlook_calendar.csv",
        "data/outlook_calendar_complete.csv",
        "outlook_calendar_complete.csv", 
        "data/outlook_calendar.csv",
        "outlook_calendar.csv"
    ]
    
    csv_path = None
    for file in csv_files:
        if os.path.exists(file):
            csv_path = file
            break
    
    if not csv_path:
        print("❌ 找不到 CSV 檔案")
        print("請先執行以下命令生成 CSV 檔案:")
        print("python3 dump_outlook_calendar.py")
        sys.exit(1)
    
    print(f"📁 使用 CSV 檔案: {csv_path}")
    
    # 檢查 Google API 憑證檔案
    client_secret_files = [
        "data/client_secret.json",
        "client_secret.json",
        "data/client_secret_454302710199-eltj3sk10l5af60aloctrvaefi891vbk.apps.googleusercontent.com.json",
        "client_secret_454302710199-eltj3sk10l5af60aloctrvaefi891vbk.apps.googleusercontent.com.json",
        "data/credentials.json",
        "credentials.json"
    ]
    
    client_secret_file = None
    for file in client_secret_files:
        if os.path.exists(file):
            client_secret_file = file
            break
    
    if not client_secret_file:
        print("❌ 找不到 Google API 憑證檔案")
        print("請從 Google Cloud Console 下載 OAuth 2.0 憑證檔案並命名為 'client_secret.json'")
        sys.exit(1)
    
    print(f"🔑 使用憑證檔案: {client_secret_file}")
    
    # 創建同步器並執行
    syncer = OutlookToGoogleCalendarSync(
        csv_path=csv_path,
        client_secret_file=client_secret_file,
        force_update=args.force,
        mark_deleted=mark_deleted,
        cleanup_days=args.cleanup_days,
        enable_cleanup=enable_cleanup
    )
    
    try:
        syncer.authenticate()
        syncer.load_cache()
        syncer.sync_events()
        
    except KeyboardInterrupt:
        print("\n⏹️  同步已中斷")
        syncer.save_cache()
    except Exception as e:
        print(f"❌ 執行錯誤: {e}")
        syncer.save_cache()

if __name__ == "__main__":
    main()
