# core/llm_translator.py
"""
LLM 翻譯模組 - 使用 Azure OpenAI 進行專業安規術語翻譯
支援 IEC 62368-1 & CNS 15598-1 標準術語
支援併發翻譯加速處理
"""
import os
import re
from typing import Optional, List, Dict
from functools import lru_cache
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

# 嘗試導入 OpenAI
try:
    from openai import AzureOpenAI
    HAS_OPENAI = True
except ImportError:
    HAS_OPENAI = False


# 系統提示詞 - 專業安規工程師角色
SYSTEM_PROMPT = """你是一位專業嚴謹的安規工程師，專精於 IEC 62368-1 與 CNS 15598-1 (109年版) 標準。
你的任務是將 CB 測試報告中的英文內容翻譯為繁體中文。

翻譯原則：
1. 使用 CNS 15598-1 官方標準術語，不可自行創造詞彙
2. 技術術語必須準確，例如：
   - SELV → 安全特低電壓
   - HAZARDOUS VOLTAGE → 危險電壓
   - BASIC INSULATION → 基本絕緣
   - SUPPLEMENTARY INSULATION → 補充絕緣
   - REINFORCED INSULATION → 加強絕緣
   - DOUBLE INSULATION → 雙重絕緣
   - PROTECTIVE EARTHING → 保護接地
   - FUNCTIONAL EARTHING → 功能接地
   - ENCLOSURE → 外殼
   - ACCESSIBLE PART → 可接觸部位
   - ENERGY SOURCE → 能量來源
   - SAFEGUARD → 安全防護
   - THERMAL CUT-OUT → 熱切斷器
   - THERMAL LINK → 熱熔斷器
   - PROTECTIVE IMPEDANCE → 保護阻抗
   - CURRENT LIMITER → 限流器
   - CREEPAGE DISTANCE → 沿面距離
   - CLEARANCE → 電氣間隙
   - WORKING VOLTAGE → 工作電壓
   - DIELECTRIC STRENGTH → 介電強度
   - TOUCH CURRENT → 接觸電流
   - PROTECTIVE CONDUCTOR CURRENT → 保護導體電流
   - FIRE ENCLOSURE → 防火外殼
   - ORDINARY PERSON → 一般人員
   - INSTRUCTED PERSON → 受指導人員
   - SKILLED PERSON → 熟練人員
3. 判定結果翻譯：
   - PASS / P → 符合
   - FAIL / F → 不符合
   - N/A → 不適用
4. 保留數值、單位、型號、標準編號（如 IEC 60950-1）不翻譯
5. 保留表格編號格式（如 Table 4.1.2 → 表 4.1.2）
6. 條款編號保持原格式（如 4.2.1、B.3）
7. 翻譯要簡潔專業，不加額外解釋

只回覆翻譯結果，不要加任何前綴或說明。"""


class LLMTranslator:
    """LLM 翻譯器 - Azure OpenAI（支援併發）"""

    def __init__(self, max_workers: int = 5):
        self.enabled = False
        self.client = None
        self.deployment = None
        self._cache: Dict[str, str] = {}
        self._cache_lock = threading.Lock()
        self.max_workers = max_workers  # 併發數量

        if not HAS_OPENAI:
            print("[LLM] openai 套件未安裝，LLM 翻譯功能停用")
            return

        # 從環境變數讀取設定
        endpoint = os.getenv("AZURE_OPENAI_ENDPOINT", "https://whaleforce-eastus2-resource.cognitiveservices.azure.com/")
        api_key = os.getenv("AZURE_OPENAI_API_KEY")
        self.deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-5.1")
        api_version = os.getenv("AZURE_OPENAI_API_VERSION", "2024-12-01-preview")

        if not api_key:
            print("[LLM] 未設定 AZURE_OPENAI_API_KEY，LLM 翻譯功能停用")
            return

        try:
            self.client = AzureOpenAI(
                api_version=api_version,
                azure_endpoint=endpoint,
                api_key=api_key,
            )
            self.enabled = True
            print(f"[LLM] Azure OpenAI 翻譯已啟用 (deployment: {self.deployment})")
        except Exception as e:
            print(f"[LLM] Azure OpenAI 初始化失敗: {e}")

    def _is_chinese(self, text: str) -> bool:
        """檢查文本是否主要為中文"""
        if not text:
            return True
        chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', text))
        total_chars = len(re.sub(r'\s+', '', text))
        if total_chars == 0:
            return True
        return chinese_chars / total_chars > 0.3

    def _should_translate(self, text: str) -> bool:
        """判斷是否需要翻譯"""
        if not text or len(text.strip()) < 3:
            return False
        # 已經是中文為主
        if self._is_chinese(text):
            return False
        # 純數字或符號
        if re.match(r'^[\d\s\.\-\+\/%°℃Ω]+$', text):
            return False
        # 標準編號（如 IEC 60950-1）
        if re.match(r'^[A-Z]+\s*\d+[\-\d\.]*$', text.strip()):
            return False
        return True

    def translate(self, text: str) -> str:
        """翻譯單個文本"""
        if not self.enabled or not self._should_translate(text):
            return text

        # 檢查快取（thread-safe）
        cache_key = text.strip()
        with self._cache_lock:
            if cache_key in self._cache:
                return self._cache[cache_key]

        try:
            response = self.client.chat.completions.create(
                model=self.deployment,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": f"翻譯以下內容：\n{text}"}
                ],
                max_completion_tokens=500,
                temperature=0.1,  # 低溫度確保一致性
            )
            result = response.choices[0].message.content.strip()
            with self._cache_lock:
                self._cache[cache_key] = result
            return result
        except Exception as e:
            print(f"[LLM] 翻譯失敗: {e}")
            return text

    def _translate_single_for_batch(self, text: str, idx: int) -> tuple:
        """併發翻譯的單個任務"""
        try:
            result = self.translate(text)
            return (idx, result)
        except Exception as e:
            print(f"[LLM] 併發翻譯失敗 (idx={idx}): {e}")
            return (idx, text)

    def translate_batch(self, texts: List[str]) -> List[str]:
        """批次翻譯多個文本（併發處理）"""
        if not self.enabled:
            return texts

        # 過濾出需要翻譯的文本
        to_translate = []
        indices = []
        results = list(texts)

        for i, text in enumerate(texts):
            if self._should_translate(text):
                cache_key = text.strip()
                with self._cache_lock:
                    if cache_key in self._cache:
                        results[i] = self._cache[cache_key]
                        continue
                to_translate.append(text)
                indices.append(i)

        if not to_translate:
            return results

        print(f"[LLM] 開始併發翻譯 {len(to_translate)} 個項目 (max_workers={self.max_workers})...")

        # 使用 ThreadPoolExecutor 併發翻譯
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            futures = {
                executor.submit(self._translate_single_for_batch, text, i): i
                for i, text in enumerate(to_translate)
            }

            completed = 0
            for future in as_completed(futures):
                try:
                    batch_idx, translated = future.result()
                    original_idx = indices[batch_idx]
                    results[original_idx] = translated
                    completed += 1
                    if completed % 10 == 0:
                        print(f"[LLM] 進度: {completed}/{len(to_translate)}")
                except Exception as e:
                    print(f"[LLM] 併發任務失敗: {e}")

        print(f"[LLM] 併發翻譯完成: {completed}/{len(to_translate)}")
        return results

    def final_review(self, texts: Dict[str, str]) -> Dict[str, str]:
        """最終審查 - 檢查並修正遺漏的英文"""
        if not self.enabled:
            return texts

        # 找出仍有大量英文的欄位
        to_review = {}
        for key, value in texts.items():
            if value and not self._is_chinese(value) and self._should_translate(value):
                to_review[key] = value

        if not to_review:
            return texts

        print(f"[LLM] 最終審查：發現 {len(to_review)} 個未翻譯欄位")

        # 批次翻譯遺漏項目
        keys = list(to_review.keys())
        values = list(to_review.values())
        translated = self.translate_batch(values)

        result = dict(texts)
        for i, key in enumerate(keys):
            result[key] = translated[i]

        return result


# 全局翻譯器實例
_translator: Optional[LLMTranslator] = None


def get_translator() -> LLMTranslator:
    """獲取全局翻譯器"""
    global _translator
    if _translator is None:
        _translator = LLMTranslator()
    return _translator


def llm_translate(text: str) -> str:
    """便捷函數：翻譯單個文本"""
    return get_translator().translate(text)


def llm_translate_batch(texts: List[str]) -> List[str]:
    """便捷函數：批次翻譯"""
    return get_translator().translate_batch(texts)


def llm_final_review(texts: Dict[str, str]) -> Dict[str, str]:
    """便捷函數：最終審查"""
    return get_translator().final_review(texts)
