#!/usr/bin/env python3
"""
IAthon
- Live pricing fetch from OpenAI with fallbacks
- Prevents file reading errors
- Generates comprehensive Markdown reports
- Creates multiple diverse visualizations
- Performs statistical tests
- Always shows cost summary
- Enhanced Streamlit UI with real-time progress
"""
import argparse
import sys
import os
import shutil      
import subprocess
import re
import time
import warnings
import json
import io
import tempfile
import contextlib
from pathlib import Path
from dataclasses import dataclass, field
from typing import Dict, Optional
from datetime import datetime
warnings.filterwarnings("ignore")

# --- Core scientific stack ---
import pandas as pd
import numpy as np
from scipy import stats

# --- Plotting (headless-safe) ---
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns  
import plotly.express as px
import plotly.graph_objects as go
from openai import OpenAI

# For web scraping (optional - will fallback if not available)
try:
    import requests
    from bs4 import BeautifulSoup
    WEB_SCRAPING_AVAILABLE = True
except ImportError:
    WEB_SCRAPING_AVAILABLE = False
    print("  requests/beautifulsoup4 not available. Install for live pricing: pip install requests beautifulsoup4")

# Try to import python-pptx
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

# ======================================================
# FALLBACK PRICING (if web fetch fails)
# Updated January 2025
# ======================================================
FALLBACK_PRICING = {
    "gpt-4o": {"input": 2.50, "output": 10.00},
    "gpt-4o-mini": {"input": 0.150, "output": 0.600},
    "o1": {"input": 15.00, "output": 60.00},
    "o1-mini": {"input": 3.00, "output": 12.00},
    "o3-mini": {"input": 1.10, "output": 4.40},
    "gpt-4-turbo": {"input": 10.00, "output": 30.00},
    "gpt-4": {"input": 30.00, "output": 60.00},
    "gpt-3.5-turbo": {"input": 0.50, "output": 1.50},
    "dall-e-3": {"standard_1024": 0.040, "standard_1792": 0.080, "hd_1024": 0.080, "hd_1792": 0.120},
    "dall-e-2": {"1024": 0.020, "512": 0.018, "256": 0.016},
}

# ======================================================
# Token Usage Tracker with Live Pricing
# ======================================================
@dataclass
class TokenUsage:
    """Track token usage per model"""
    prompt_tokens: int = 0
    completion_tokens: int = 0
    total_tokens: int = 0
    
    def add(self, prompt: int, completion: int):
        self.prompt_tokens += prompt
        self.completion_tokens += completion
        self.total_tokens += (prompt + completion)

@dataclass
class UsageTracker:
    """Central tracker with live pricing fetch"""
    models: Dict[str, TokenUsage] = field(default_factory=dict)
    pricing_cache: Dict[str, Dict[str, float]] = field(default_factory=dict)
    pricing_source: str = "unknown"
    cache_file: Path = field(default_factory=lambda: Path.home() / ".cache" / "iathlon" / "openai_pricing.json")
    
    def __post_init__(self):
        self.cache_file.parent.mkdir(parents=True, exist_ok=True)
        self.fetch_pricing()
    
    def _fetch_live_pricing_from_web(self) -> Optional[Dict]:
        """Attempt to fetch pricing from OpenAI's website"""
        if not WEB_SCRAPING_AVAILABLE:
            return None
        
        print("📡 Fetching live pricing from OpenAI...")
        
        try:
            response = requests.get(
                "https://openai.com/api/pricing/",
                timeout=15,
                headers={
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                }
            )
            
            if response.status_code == 200:
                text = response.text
                extracted = {}
                
                # Try to find JSON pricing data
                json_patterns = [r'"pricing":\s*({[^}]+})', r'pricing:\s*({[^}]+})']
                for pattern in json_patterns:
                    for match in re.finditer(pattern, text):
                        try:
                            pricing_json = json.loads(match.group(1))
                            if isinstance(pricing_json, dict):
                                extracted.update(pricing_json)
                        except:
                            continue
                
                # Fallback: regex patterns
                price_patterns = [
                    r'(gpt-[\w.-]+|o\d+(?:-[\w]+)?)\s*[^\d]*\$\s*([\d.]+)\s*(?:/|per)\s*(?:1M|million)[^\$]*\$\s*([\d.]+)\s*(?:/|per)\s*(?:1M|million)',
                ]
                
                for pattern in price_patterns:
                    for match in re.finditer(pattern, text, re.IGNORECASE):
                        try:
                            model = match.group(1).strip().lower()
                            input_price = float(match.group(2))
                            output_price = float(match.group(3))
                            extracted[model] = {"input": input_price, "output": output_price}
                        except:
                            continue
                
                if extracted:
                    print(f"✅ Fetched pricing for {len(extracted)} models")
                    cache_data = {
                        'timestamp': datetime.now().isoformat(),
                        'source': 'https://openai.com/api/pricing/',
                        'pricing': extracted
                    }
                    with open(self.cache_file, 'w') as f:
                        json.dump(cache_data, f, indent=2)
                    return extracted
                else:
                    print(f"⚠️ Page loaded but no pricing extracted (JavaScript-rendered content)")
        except Exception as e:
            print(f"❌ Web fetch failed: {str(e)[:80]}")
        
        return None
    
    def _load_from_cache(self) -> Optional[Dict]:
        """Load from cache"""
        try:
            if self.cache_file.exists():
                with open(self.cache_file, 'r') as f:
                    cache_data = json.load(f)
                timestamp = cache_data.get('timestamp', 'unknown')
                pricing = cache_data.get('pricing', {})
                if pricing:
                    print(f"💾 Loaded cached pricing from {timestamp}")
                    return pricing
        except Exception as e:
            print(f"⚠️ Cache load failed: {e}")
        return None
    
    def fetch_pricing(self):
        """Fetch with fallback chain"""
        live_pricing = self._fetch_live_pricing_from_web()
        if live_pricing:
            self.pricing_cache = live_pricing
            self.pricing_source = "live"
            return
        
        cached_pricing = self._load_from_cache()
        if cached_pricing:
            self.pricing_cache = cached_pricing
            self.pricing_source = "cache"
            return
        
        print("📋 Using fallback pricing (January 2025)")
        self.pricing_cache = FALLBACK_PRICING.copy()
        self.pricing_source = "fallback"
    
    def _get_model_pricing(self, model_name: str) -> Dict[str, float]:
        model_lower = model_name.lower()
        if model_name in self.pricing_cache:
            return self.pricing_cache[model_name]
        
        # Priority matching for known tiers
        if "o1" in model_lower or "gpt-4" in model_lower and "mini" not in model_lower:
            return FALLBACK_PRICING.get("gpt-4o", {"input": 2.50, "output": 10.00})
        if "mini" in model_lower or "gpt-3.5" in model_lower:
            return FALLBACK_PRICING.get("gpt-4o-mini", {"input": 0.15, "output": 0.60})
            
        return {"input": 2.50, "output": 10.00}
        
    def record(self, model: str, usage):
        """Record usage from OpenAI response"""
        if model not in self.models:
            self.models[model] = TokenUsage()
        
        prompt = 0
        completion = 0
        
        for prompt_attr in ['prompt_tokens', 'input_tokens', 'total_input_tokens']:
            if hasattr(usage, prompt_attr):
                prompt = getattr(usage, prompt_attr)
                break
        
        for completion_attr in ['completion_tokens', 'output_tokens', 'total_output_tokens']:
            if hasattr(usage, completion_attr):
                completion = getattr(usage, completion_attr)
                break
        
        if prompt == 0 and completion == 0:
            try:
                if hasattr(usage, '__dict__'):
                    usage_dict = usage.__dict__
                elif hasattr(usage, 'model_dump'):
                    usage_dict = usage.model_dump()
                else:
                    usage_dict = dict(usage)
                prompt = usage_dict.get('prompt_tokens', 0) or usage_dict.get('input_tokens', 0)
                completion = usage_dict.get('completion_tokens', 0) or usage_dict.get('output_tokens', 0)
            except:
                pass
        
        self.models[model].add(prompt, completion)
    
    def record_image(self, model: str, num_images: int = 1, size: str = "1024x1024", quality: str = "standard"):
        """Record image generation"""
        if model not in self.models:
            self.models[model] = TokenUsage()
        self.models[model].add(num_images, 0)
        if not hasattr(self.models[model], 'image_specs'):
            self.models[model].image_specs = []
        self.models[model].image_specs.append({'size': size, 'quality': quality})
    
    def print_summary(self):
        """Print cost summary"""
        print("\n" + "="*80)
        print("💰 TOKEN USAGE & COST SUMMARY")
        print("="*80)
        
        source_indicators = {
            'live': '🌐 Live from OpenAI',
            'cache': '💾 Cached pricing',
            'fallback': '📋 Fallback pricing'
        }
        print(f"Pricing: {source_indicators.get(self.pricing_source, 'Unknown')}")
        print(f"⏰ {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("="*80)
        
        total_cost = 0.0
        
        for model, usage in self.models.items():
            print(f"\n📊 {model}")
            
            if 'dall-e' in model.lower():
                num_images = usage.prompt_tokens
                print(f"   Images: {num_images}")
                if hasattr(usage, 'image_specs'):
                    model_cost = 0
                    pricing = self._get_model_pricing(model)
                    for spec in usage.image_specs:
                        size, quality = spec['size'], spec['quality']
                        if 'dall-e-3' in model.lower():
                            key = f"{'hd' if quality == 'hd' else 'standard'}_{size.split('x')[0]}"
                            cost = pricing.get(key, 0.040)
                        else:
                            cost = pricing.get(size.split('x')[0], 0.020)
                        model_cost += cost
                    total_cost += model_cost
                    print(f"   Cost: ${model_cost:.6f}")
            else:
                print(f"   Input:  {usage.prompt_tokens:>10,} tokens")
                print(f"   Output: {usage.completion_tokens:>10,} tokens")
                print(f"   Total:  {usage.total_tokens:>10,} tokens")
                
                pricing = self._get_model_pricing(model)
                input_cost = (usage.prompt_tokens / 1_000_000) * pricing["input"]
                output_cost = (usage.completion_tokens / 1_000_000) * pricing["output"]
                model_cost = input_cost + output_cost
                total_cost += model_cost
                
                print(f"   Rate: ${pricing['input']:.3f}/${pricing['output']:.3f} per 1M")
                print(f"   Cost: ${model_cost:.6f}")
        
        print(f"\n{'='*80}")
        print(f"💵 TOTAL COST: ${total_cost:.6f}")
        print("="*80)
        
        return total_cost
    
    def save_log(self, output_dir: Path):
        """Save log file"""
        log_path = output_dir / "token_usage.log"
        with open(log_path, 'w') as f:
            f.write(f"TOKEN USAGE LOG\n")
            f.write(f"Generated: {datetime.now()}\n")
            f.write(f"Pricing: {self.pricing_source}\n\n")
            
            total = 0.0
            for model, usage in self.models.items():
                f.write(f"{model}:\n")
                if 'dall-e' not in model.lower():
                    pricing = self._get_model_pricing(model)
                    cost = (usage.prompt_tokens / 1_000_000) * pricing["input"] + \
                           (usage.completion_tokens / 1_000_000) * pricing["output"]
                    f.write(f"  Tokens: {usage.total_tokens:,}\n")
                    f.write(f"  Cost: ${cost:.6f}\n")
                    total += cost
                f.write("\n")
            f.write(f"TOTAL: ${total:.6f}\n")
        print(f"📄 Log saved: {log_path}")

# ======================================================
# Utilities
# ======================================================
def get_api_key(cli_key=None):
    return cli_key or os.getenv("OPENAI_API_KEY")

def robust_rmtree(path: Path):
    """Delete directory with retry"""
    if not path.exists():
        return
    for _ in range(3):
        try:
            shutil.rmtree(path)
            return
        except PermissionError:
            time.sleep(1)
    for item in path.iterdir():
        try:
            if item.is_file(): 
                item.unlink()
            elif item.is_dir(): 
                shutil.rmtree(item)
        except:
            pass

ALLOWED_KINDS = {"histogram", "correlation_heatmap", "boxplot", "scatter"}

def validate_plan(plan: dict, num_features: list, cat_features: list):
    """Validate plan structure and content"""
    if not isinstance(plan, dict):
        raise ValueError("plan must be a dictionary")
    if "stop" not in plan or "language" not in plan or "analyses" not in plan:
        raise ValueError("plan incomplete: missing stop/language/analyses keys")
    
    analyses = plan["analyses"]
    if not isinstance(analyses, list):
        raise ValueError("plan['analyses'] must be a list")
    
    if plan.get("stop") is False and len(analyses) == 0:
        raise ValueError("If stop=false, must have at least 1 analysis")
    
    if len(analyses) > 10:
        raise ValueError(f"Too many analyses requested: {len(analyses)} (max 10)")
    
    all_cols = set(num_features) | set(cat_features)
    stats_list = plan.get("stats", [])
    if stats_list is None:
        stats_list = []
    if not isinstance(stats_list, list):
        raise ValueError("plan['stats'] must be a list if provided")
    
    seen_filenames = set()
    for i, a in enumerate(analyses):
        if not isinstance(a, dict):
            raise ValueError(f"analysis[{i}] must be a dictionary")
        
        kind = a.get("kind")
        if kind not in ALLOWED_KINDS:
            raise ValueError(f"analysis[{i}].kind invalid: {kind}. Must be one of {ALLOWED_KINDS}")
        
        fn = a.get("filename")
        if not isinstance(fn, str) or not fn.endswith(".png"):
            raise ValueError(f"analysis[{i}].filename invalid: {fn} (must end with .png)")
        
        if fn in seen_filenames:
            raise ValueError(f"duplicate filename: {fn}")
        seen_filenames.add(fn)
        
        # Requirements per kind
        if kind == "histogram":
            x = a.get("x")
            if not x or x not in all_cols:
                raise ValueError(f"histogram requires valid x column. got: {x}")
        
        if kind == "scatter":
            x, y = a.get("x"), a.get("y")
            if not x or x not in all_cols or not y or y not in all_cols:
                raise ValueError(f"scatter requires valid x,y columns. got: x={x}, y={y}")
        
        if kind == "boxplot":
            y = a.get("y")
            if not y or y not in num_features:
                raise ValueError(f"boxplot requires numeric y from num_features. got: {y}")
            for k in ["x", "hue"]:
                if a.get(k) and a.get(k) not in all_cols:
                    raise ValueError(f"boxplot {k} invalid: {a.get(k)}")
    
    # Soft validation for stats tests
    for i, s in enumerate(stats_list):
        if not isinstance(s, dict):
            raise ValueError(f"stats[{i}] must be a dictionary")
        test = s.get("test")
        if not isinstance(test, str):
            raise ValueError(f"stats[{i}].test must be a string")
        for key in ["x", "y", "group"]:
            if s.get(key) and s.get(key) not in all_cols:
                raise ValueError(f"stats[{i}].{key} invalid: {s.get(key)}")

# ======================================================
# Plan Sanitization
# ======================================================
def sanitize_plan(plan: dict, num_features: list, cat_features: list) -> dict:
    """Repair common LLM plan mistakes to avoid aborting."""
    if not isinstance(plan, dict):
        return plan
    
    analyses = plan.get("analyses", [])
    if not isinstance(analyses, list):
        return plan
    
    default_num = num_features[0] if num_features else None
    default_cat = cat_features[0] if cat_features else None
    allowed = {"histogram", "correlation_heatmap", "boxplot", "scatter"}
    
    for a in analyses:
        if not isinstance(a, dict):
            continue
        
        # Normalize kind
        kind = a.get("kind")
        if isinstance(kind, str):
            kind = kind.strip().lower()
            a["kind"] = kind
        
        # Fix common invalid kind
        if a.get("kind") == "violin":
            a["kind"] = "boxplot"
        
        # Drop completely unknown kinds
        if a.get("kind") not in allowed:
            a["__drop__"] = True
            continue
        
        # Ensure required fields
        if a["kind"] == "boxplot":
            if not a.get("y") and default_num:
                a["y"] = default_num
            if not a.get("x") and not a.get("hue") and default_cat:
                a["x"] = default_cat
        
        if a["kind"] == "histogram":
            if not a.get("x") and default_num:
                a["x"] = default_num
        
        if a["kind"] == "scatter":
            if not a.get("x") and default_num:
                a["x"] = default_num
            if not a.get("y") and len(num_features) > 1:
                a["y"] = num_features[1]
            elif not a.get("y") and default_num:
                a["y"] = default_num
    
    # Remove dropped analyses
    plan["analyses"] = [a for a in analyses if not a.get("__drop__")]
    
    # If stop=false but nothing usable remains, stop gracefully
    if plan.get("stop") is False and len(plan["analyses"]) == 0:
        plan["stop"] = True
    
    return plan

# ======================================================
# STAT Lines Parsing
# ======================================================
def parse_stat_lines(stdout_text: str):
    lines = []
    if not stdout_text:
        return lines
    for ln in stdout_text.splitlines():
        ln = ln.strip()
        if ln.startswith("STAT:"):
            lines.append(ln)
    return lines

# ======================================================
# Data Loading
# ======================================================
def load_and_clean(path: Path):
    suffix = path.suffix.lower()
    try:
        # 1. CARGA ROBUSTA
        if suffix == ".csv":
            df = pd.read_csv(path, sep=None, engine='python', encoding='utf-8-sig', on_bad_lines='skip')
        else:
            df = pd.read_excel(path)

        # 2. LIMPIEZA PROFUNDA
        df.columns = [str(c).strip().replace('.', '_').replace(' ', '_') for c in df.columns]
        df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
        df = df.replace(['None', 'Unknown', 'nan', 'NULL', ''], np.nan)

        # 3. TRATAMIENTO DE TIEMPOS
        for col in df.select_dtypes(include=['object']).columns:
            sample = df[col].dropna().head(10).astype(str)
            if len(sample) > 0 and sample.str.contains(r'^(\d+:)?\d+:\d+(\.\d+)?$').all():
                df[col] = pd.to_timedelta(df[col], errors='coerce').dt.total_seconds()
                print(f"⏱️  Columna de tiempo convertida a segundos: {col}")

        # 4. DETECCIÓN DE IDs
        for col in df.select_dtypes(include=np.number).columns:
            if 'id' in col.lower() or 'bib' in col.lower() or (df[col].nunique() == len(df) and len(df) > 100):
                df[col] = df[col].astype(str)
                print(f"🆔 Detectada columna de identidad (excluida de análisis numérico): {col}")

        # 5. FILTRADO DE COLUMNAS
        df = df.dropna(axis=1, thresh=len(df) * 0.1)
        df = df.loc[:, df.nunique() > 1]

        num_features = df.select_dtypes(include=np.number).columns.tolist()
        cat_features = df.select_dtypes(include=["object", "category", "datetime64"]).columns.tolist()

        return df, num_features, cat_features
    except Exception as e:
        print(f"❌ Error al cargar {path.name}: {e}")
        raise

# ======================================================
# Safe Execution Environment
# ======================================================
class SafeRunner:
    def __init__(self, df, num_features, cat_features, output_dir: Path, verbose=False):
        self.df = df
        self.num_features = num_features
        self.cat_features = cat_features
        self.output_dir = Path(output_dir).absolute()
        
        if self.output_dir.name.lower() == "figures":
            self.figures_dir = self.output_dir
        else:
            self.figures_dir = self.output_dir / "figures"
            
        self.figures_dir.mkdir(parents=True, exist_ok=True)
        self.verbose = verbose
            
    def run(self, code, plan=None):
        """Execute generated code in safe environment"""
        assert "exec(" not in code.lower(), "Generated code must not call exec()"
        
        # Extract code from markdown if present
        if "```" in code:
            match = re.search(r"```(?:python)?\n?(.*?)\n?```", code, re.DOTALL)
            if match:
                code = match.group(1)
            else:
                code = "\n".join([line for line in code.split("\n") if "```" not in line])
        
        # Block file reading operations
        forbidden = [
            (r'pd\.read_csv\s*\(', "pd.read_csv()"),
            (r'pd\.read_excel\s*\(', "pd.read_excel()"),
            (r'pd\.read_json\s*\(', "pd.read_json()"),
        ]
        
        for pattern, func in forbidden:
            if re.search(pattern, code, re.IGNORECASE):
                raise RuntimeError(f"❌ Attempted {func} - use existing 'df' variable instead")
        
        df_copy = self.df.copy()
        
        # Safe globals environment
        globals_safe = {
            "df": df_copy,
            "num_features": self.num_features, 
            "cat_features": self.cat_features, 
            "plt": plt, 
            "sns": sns, 
            "px": px, 
            "go": go, 
            "stats": stats, 
            "np": np, 
            "pd": pd,
            "FIGURES_DIR": str(self.figures_dir),
            "os": os,
            "plan": plan
        }
        
        # Strip 'figures/' prefix to prevent double-nesting
        abs_fig_path = str(self.figures_dir).replace('\\', '/')
        code = re.sub(
            r"['\"](?:figures/)?([^'\"\s]+\.png)['\"]", 
            f"'{abs_fig_path}/\\1'", 
            code
        )
        
        old_cwd = os.getcwd()
        os.chdir(self.figures_dir)
        buf = io.StringIO()
        
        try:
            with contextlib.redirect_stdout(buf):
                exec(code, globals_safe)
        except NameError as e:
            raise RuntimeError(
                f"Generated code referenced an undefined variable: {e}. "
                "Ensure all intermediate variables are defined before being used."
            )
        except Exception as e:
            raise RuntimeError(str(e))
        finally:
            os.chdir(old_cwd)
        
        stdout_text = buf.getvalue()
        pngs = list(self.figures_dir.glob("*.png"))
        
        if not pngs:
            raise RuntimeError("No figures were generated. Ensure plots are saved to FIGURES_DIR.")
        
        return stdout_text

# ======================================================
# Data Quality Validator
# ======================================================
class DataQualityValidator:
    def __init__(self, df, num_features, cat_features, verbose=False):
        self.df = df
        self.num_features = num_features
        self.cat_features = cat_features
        self.verbose = verbose
        self.issues = []
        self.corrections = []
        
    def validate(self):
        """Run all validation checks"""
        print("\n" + "="*80)
        print("🔍 DATA QUALITY VALIDATION")
        print("="*80)
        
        self._check_synthetic_patterns()
        self._check_outliers()
        self._check_domain_constraints()
        self._check_temporal_logic()
        
        return self._generate_report()
    
    def _check_synthetic_patterns(self):
        """Detect artificially generated data"""
        print("\n🤖 Checking for synthetic data patterns...")
        
        for col in self.num_features:
            data = self.df[col].dropna()
            if len(data) < 30:
                continue
            
            # Test 1: Too-perfect uniform distribution
            _, p_uniform = stats.kstest(data, 'uniform', 
                                       args=(data.min(), data.max() - data.min()))
            if p_uniform > 0.95:
                self.issues.append({
                    'type': 'SYNTHETIC_UNIFORM',
                    'column': col,
                    'severity': 'WARNING',
                    'message': f'{col} follows suspiciously perfect uniform distribution (p={p_uniform:.4f})',
                    'recommendation': 'Verify data source - may be artificially generated'
                })
            
            # Test 2: Too-perfect normal distribution
            _, p_normal = stats.normaltest(data)
            if p_normal > 0.99:
                self.issues.append({
                    'type': 'SYNTHETIC_NORMAL',
                    'column': col,
                    'severity': 'WARNING',
                    'message': f'{col} follows suspiciously perfect normal distribution (p={p_normal:.4f})',
                    'recommendation': 'Real-world data rarely this normal - verify authenticity'
                })
            
            # Test 3: Repeated patterns
            value_counts = data.value_counts()
            if len(value_counts) < len(data) * 0.1 and len(data) > 100:
                self.issues.append({
                    'type': 'EXCESSIVE_DUPLICATES',
                    'column': col,
                    'severity': 'WARNING',
                    'message': f'{col} has only {len(value_counts)} unique values from {len(data)} records',
                    'recommendation': 'Verify if data was artificially constrained'
                })
            
            # Test 4: Too many round numbers
            if data.dtype in [np.float64, np.float32]:
                round_nums = (data == data.round(0)).sum()
                if round_nums / len(data) > 0.9 and len(data) > 50:
                    self.issues.append({
                        'type': 'EXCESSIVE_ROUNDING',
                        'column': col,
                        'severity': 'INFO',
                        'message': f'{col} has {round_nums/len(data)*100:.1f}% round numbers',
                        'recommendation': 'Natural measurements usually have decimal precision'
                    })
    
    def _check_outliers(self):
        """Detect statistical outliers"""
        print("📊 Checking for outliers...")
        
        for col in self.num_features:
            data = self.df[col].dropna()
            if len(data) < 10:
                continue
            
            # IQR method
            Q1 = data.quantile(0.25)
            Q3 = data.quantile(0.75)
            IQR = Q3 - Q1
            lower_bound = Q1 - 3 * IQR
            upper_bound = Q3 + 3 * IQR
            
            outliers = data[(data < lower_bound) | (data > upper_bound)]
            
            if len(outliers) > 0:
                self.issues.append({
                    'type': 'OUTLIERS',
                    'column': col,
                    'severity': 'INFO',
                    'message': f'{col} has {len(outliers)} extreme outliers',
                    'details': f'Range: [{data.min():.2f}, {data.max():.2f}], Expected: [{lower_bound:.2f}, {upper_bound:.2f}]',
                    'recommendation': 'Review extreme values - may indicate data entry errors'
                })
    
    def _check_domain_constraints(self):
        """Check domain-specific logical constraints"""
        print("🔬 Checking domain constraints...")
        
        # Age checks
        age_cols = [c for c in self.num_features if 'age' in c.lower()]
        for col in age_cols:
            invalid_ages = self.df[(self.df[col] < 0) | (self.df[col] > 120)]
            if len(invalid_ages) > 0:
                self.issues.append({
                    'type': 'INVALID_AGE',
                    'column': col,
                    'severity': 'ERROR',
                    'message': f'{col} has {len(invalid_ages)} invalid ages (< 0 or > 120)',
                    'details': f'Invalid values: {invalid_ages[col].tolist()[:5]}',
                    'recommendation': 'CRITICAL: Fix or remove invalid ages'
                })
        
        # Percentage checks
        pct_cols = [c for c in self.num_features if any(x in c.lower() for x in ['percent', 'pct', 'rate'])]
        for col in pct_cols:
            invalid_pct = self.df[(self.df[col] < 0) | (self.df[col] > 100)]
            if len(invalid_pct) > 0:
                self.issues.append({
                    'type': 'INVALID_PERCENTAGE',
                    'column': col,
                    'severity': 'ERROR',
                    'message': f'{col} has {len(invalid_pct)} values outside 0-100%',
                    'recommendation': 'Fix percentage values'
                })
        
        # Time/duration checks
        time_cols = [c for c in self.num_features if any(x in c.lower() for x in ['time', 'duration', 'seconds', 'minutes', 'hours'])]
        for col in time_cols:
            negative_times = self.df[self.df[col] < 0]
            if len(negative_times) > 0:
                self.issues.append({
                    'type': 'NEGATIVE_TIME',
                    'column': col,
                    'severity': 'ERROR',
                    'message': f'{col} has {len(negative_times)} negative time values',
                    'details': f'Examples: {negative_times[col].tolist()[:5]}',
                    'recommendation': 'CRITICAL: Times cannot be negative'
                })
    
    def _check_temporal_logic(self):
        """Check temporal relationships"""
        print("⏰ Checking temporal logic...")
        
        # Check for date columns
        date_cols = self.df.select_dtypes(include=['datetime64']).columns.tolist()
        date_cols += [c for c in self.df.columns if any(x in c.lower() for x in ['date', 'year', 'born', 'death'])]
        
        # Birth/death logic
        if 'birth' in str(date_cols).lower() and 'death' in str(date_cols).lower():
            birth_col = [c for c in date_cols if 'birth' in c.lower()][0]
            death_col = [c for c in date_cols if 'death' in c.lower()][0]
            
            invalid = self.df[self.df[death_col] < self.df[birth_col]]
            if len(invalid) > 0:
                self.issues.append({
                    'type': 'TEMPORAL_PARADOX',
                    'column': f'{birth_col}, {death_col}',
                    'severity': 'ERROR',
                    'message': f'{len(invalid)} records have death before birth',
                    'recommendation': 'CRITICAL: Fix temporal logic errors'
                })
    
    def _generate_report(self):
        """Generate validation report"""
        print("\n" + "="*80)
        print("📋 VALIDATION SUMMARY")
        print("="*80)
        
        if not self.issues:
            print("✅ No major data quality issues detected")
            return True
        
        # Group by severity
        errors = [i for i in self.issues if i['severity'] == 'ERROR']
        warnings = [i for i in self.issues if i['severity'] == 'WARNING']
        info = [i for i in self.issues if i['severity'] == 'INFO']
        
        if errors:
            print(f"\n🚨 {len(errors)} CRITICAL ERRORS:")
            for issue in errors:
                print(f"   • {issue['column']}: {issue['message']}")
                print(f"     → {issue['recommendation']}")
        
        if warnings:
            print(f"\n⚠️  {len(warnings)} WARNINGS:")
            for issue in warnings:
                print(f"   • {issue['column']}: {issue['message']}")
                print(f"     → {issue['recommendation']}")
        
        if info:
            print(f"\nℹ️  {len(info)} INFORMATIONAL:")
            for issue in info:
                print(f"   • {issue['column']}: {issue['message']}")
        
        print("="*80)
        return len(errors) == 0
    
    def auto_fix(self):
        """Attempt automatic fixes for common issues"""
        print("\n🔧 ATTEMPTING AUTO-FIXES...")
        
        df_fixed = self.df.copy()
        
        for issue in self.issues:
            if issue['type'] == 'NEGATIVE_TIME' and issue['severity'] == 'ERROR':
                col = issue['column']
                before_count = (df_fixed[col] < 0).sum()
                df_fixed[col] = df_fixed[col].abs()
                self.corrections.append(f"Fixed {before_count} negative times in '{col}' (took absolute value)")
                print(f"   ✅ Fixed {before_count} negative values in '{col}'")
            
            elif issue['type'] == 'INVALID_AGE':
                col = issue['column']
                before_count = ((df_fixed[col] < 0) | (df_fixed[col] > 120)).sum()
                df_fixed[col] = df_fixed[col].clip(0, 120)
                self.corrections.append(f"Capped {before_count} invalid ages in '{col}' to 0-120 range")
                print(f"   ✅ Capped {before_count} invalid ages in '{col}'")
        
        if self.corrections:
            print(f"\n✨ Applied {len(self.corrections)} automatic fixes")
            return df_fixed
        else:
            print("   ℹ️  No auto-fixes applied")
            return df_fixed

# ======================================================
# Decision Maker
# ======================================================
class DecisionMaker:
    def __init__(self, api_key, tracker, custom_prompt=None, lang="es", verbose=False):
        self.client = OpenAI(api_key=api_key)
        self.tracker = tracker
        self.custom_prompt = custom_prompt
        self.lang = lang
        self.verbose = verbose
        self.model = "gpt-4o-mini"
    
    def decide(self, observations: str, step_num: int, max_steps: int, 
            num_features: list, cat_features: list, data_preview: str) -> str:
        
        context = ""
        if self.custom_prompt:
            context = f"### OBLIGATORY DATA DESCRIPTION & BUSINESS RULES ###\n{self.custom_prompt}\n### END OF RULES ###\n\n"

        # First step: Generate comprehensive batch plan
        if step_num == 0:
            prompt = f"""{context}DATA PREVIEW:
{data_preview}

NUMERIC: {num_features}
CATEGORICAL: {cat_features}

STEP {step_num + 1}/{max_steps} - INITIAL BATCH ANALYSIS PLAN

**CRITICAL: You MUST respond with ONLY a valid JSON object. No explanations, no markdown, no text before or after.**

Required JSON structure:
{{
"stop": false,
"language": "{self.lang}",
"analyses": [
    {{"kind": "histogram", "x": "column_name", "filename": "hist_01.png", "title": "Title", "caption": "Description"}},
    {{"kind": "correlation_heatmap", "filename": "corr_01.png", "title": "Title", "caption": "Description"}},
    {{"kind": "scatter", "x": "col1", "y": "col2", "filename": "scatter_01.png", "title": "Title", "caption": "Description"}}
],
"stats": []
}}

Generate 5-7 different analyses covering:
- Distribution analysis (histograms)
- Correlation heatmap
- Relationship plots (scatter)
- Group comparisons (boxplot)

Use ONLY columns from: {num_features + cat_features}"""

        # Follow-up steps
        else:
            prompt = f"""{context}DATA PREVIEW:
{data_preview}

NUMERIC: {num_features}
CATEGORICAL: {cat_features}

STEP {step_num + 1}/{max_steps}
CURRENT OBSERVATIONS: {observations}

**CRITICAL: You MUST respond with ONLY a valid JSON object. No explanations, no markdown.**

If analysis is complete, respond with:
{{
"stop": true,
"language": "{self.lang}",
"analyses": [],
"stats": []
}}

If more analysis needed, respond with:
{{
"stop": false,
"language": "{self.lang}",
"analyses": [
    {{"kind": "boxplot", "y": "numeric_col", "x": "cat_col", "filename": "box_02.png", "title": "Title", "caption": "Description"}}
],
"stats": [
    {{"test": "ttest", "y": "numeric_col", "group": "cat_col"}}
]
}}

Available analysis kinds: histogram, correlation_heatmap, boxplot, scatter
Available statistical tests: pearson, ttest, anova, chi-square"""

        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": "You are a JSON generation engine. You ONLY output valid JSON objects. Never include markdown code blocks, explanations, or any text outside the JSON structure."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
        )
        
        if hasattr(response, 'usage'):
            self.tracker.record(self.model, response.usage)
        
        raw_response = response.choices[0].message.content.strip()
        
        # Clean up common LLM formatting issues
        if raw_response.startswith("```"):
            raw_response = re.sub(r'^```(?:json)?\n?', '', raw_response)
            raw_response = re.sub(r'\n?```$', '', raw_response)
        
        return raw_response.strip()

    def synthesize_report(self, analysis_log, original_prompt, figures_dir, data_preview, lang="es"):
        """Generate final markdown report"""
        if self.verbose:
            print("📝 Synthesizing Final Report...")
        
        # Get actual files on disk
        figs = sorted(figures_dir.glob("*.png"))
        fig_list = "\n".join([f"figures/{f.name}" for f in figs])
        
        context = f"USER CONTEXT: {self.custom_prompt}\n\n" if self.custom_prompt else ""
        analysis_log_json = json.dumps(analysis_log, ensure_ascii=False, indent=2)
        
        lang_line = "Write the report in Spanish." if lang == "es" else "Write the report in English."
        prompt = f"""{context}
You are a technical reporting engine. Write a professional, publication-ready Data Analysis
Report in Markdown format.
{lang_line}

DATA PREVIEW:
{data_preview}

ANALYSIS LOG (JSON):
{analysis_log_json}

AVAILABLE FIGURES (USE ONLY THESE FILENAMES - DO NOT DUPLICATE):
{fig_list}

FIGURE CAPTIONS RULE:
- If analysis_log contains plan.analyses items with "caption", use those captions verbatim.
- DO NOT invent captions when a caption is provided in the log.
- Use analysis_log[*].stat_lines to extract p-values, means, std, and test results.
- CRITICAL: Each figure should appear ONLY ONCE in the report. Do not duplicate figure references.

STRICT REQUIREMENTS:
1. DO NOT include conversational filler.
2. DO NOT offer to re-run analyses.
3. INTEGRATE figures: Use ONLY the filenames listed in "AVAILABLE FIGURES".
   Each figure should appear exactly ONCE in the document.
   Place each ![Description](figures/filename.png) tag immediately after the paragraph discussing that result.
4. Use professional headers (#, ##, ###).
5. Include a "Statistical Summary" section with exact p-values and means.
6. END the report with a "Conclusions" section.
7. Every table caption MUST be followed by TWO empty lines.
8. IMAGE SIZING: Use Pandoc attributes: ![Description](figures/filename.png){{width=75%}}
9. TEXT ANCHORING: Ensure every image is preceded by 3-4 substantive sentences.
10. FIGURE CAPTIONS: Every image MUST have a detailed caption below it.
    Example:
    ![Analysis of scores](figures/plot_001.png){{width=75%}}
    **Figure 1: Distribution of math scores by socioeconomic status.**
11. NO DUPLICATES: Each figure filename should appear exactly once in the entire report.

REPORT STRUCTURE:
# [Title]
## Executive Summary
## Methodology
## Analysis & Results (Embed each figure exactly once here)
## Statistical Tables
## Conclusions and Recommendations
"""
        
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": "You are a professional data science reporter. Output only the report text. Each figure should appear exactly once in the report."},
                {"role": "user", "content": prompt}
            ],
        )
        
        if hasattr(response, "usage"):
            self.tracker.record(self.model, response.usage)
        
        report_content = response.choices[0].message.content
        
        # Verify and fix duplicate figure references
        fig_names = [f.name for f in figs]
        for fig_name in fig_names:
            # Count occurrences
            pattern = f"figures/{fig_name}"
            count = report_content.count(pattern)
            if count > 1:
                print(f"⚠️  Warning: Figure {fig_name} appears {count} times. Keeping first occurrence only.")
                # Keep only the first occurrence
                parts = report_content.split(pattern)
                report_content = parts[0] + pattern + "".join(parts[1:]).replace(pattern, "")
        
        return report_content

    def generate_title_image_prompt(
        self,
        data_preview: str,
        num_features: list,
        cat_features: list,
        dataset_name: str = ""
    ) -> str:
        """
        Crea una query corta y visual (tipo Pexels) basada en preview + columnas.
        Devuelve un string de una sola línea.
        """
        context = ""
        if self.custom_prompt:
            context = f"CONTEXTO/DOMINIO:\n{self.custom_prompt}\n\n"

        prompt = f"""
{context}
Genera una consulta corta para buscar una FOTO de stock (Pexels).
Devuelve SOLO la query final (sin comillas, sin listas).

Reglas:
- Máx 12 palabras
- Debe ser una escena física y visual
- Evita palabras abstractas como: performance, analysis, characteristics, optimization
- Sin comas

Dataset name: {dataset_name}

Dataset preview:
{data_preview}

Numeric columns: {num_features}
Categorical columns: {cat_features}

Ejemplos buenos:
students studying with laptops in classroom
doctors reviewing patient charts in hospital corridor
factory workers monitoring machines on production line
athletes training outdoors with stopwatch

Ahora produce UNA única query:
""".strip()

        response = self.client.chat.completions.create(
            model=self.model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.4
        )

        if hasattr(response, "usage"):
            self.tracker.record(self.model, response.usage)

        return response.choices[0].message.content.strip()

# ======================================================
# Code Generator
# ======================================================
class CodexGenerator:
    def __init__(self, api_key, tracker: UsageTracker, verbose=False):
        self.client = OpenAI(api_key=api_key)
        self.tracker = tracker
        self.verbose = verbose
        self.model = "gpt-4o-mini"
    
    def generate(self, plan: dict, num_features, cat_features, data_card: str, error_context=None) -> str:
        """Generate Python code from analysis plan"""
        error_block = ""
        if error_context:
            error_block = f"""
PREVIOUS EXECUTION ERROR (FIX THIS):
{error_context}
RULES:
- Do NOT repeat the same mistake
- Ensure all variables are defined before use
"""
        
        plan_json = json.dumps(plan, ensure_ascii=False, indent=2)
        
        prompt = f"""You are an expert data scientist. Generate Python code.

DATACARD:
{data_card}

PLAN (JSON):
{plan_json}

{error_block}

STRICT CONTRACT:
- DO NOT modify filenames: use EXACTLY analysis["filename"]
- DO NOT generate extra plots outside of plan["analyses"]
- 1 analysis => 1 saved figure
- If a column doesn't exist or is unusable:
  - print: STAT: skipped <kind> reason=<...>
  - DO NOT crash; continue with next analysis
- After saving each figure: plt.close() ALWAYS
- After saving each figure: print a STAT line with key metrics
- For each stats test in plan["stats"]: compute and print a STAT line
  - If a test cannot be run: print STAT: skipped test=<...> reason=<...>

RULES:
- df already exists, DO NOT read files
- Save each plot to: os.path.join(FIGURES_DIR, filename_from_plan)
- Use plt.savefig(..., dpi=150, bbox_inches='tight')
- Use ONLY existing columns
- Titles/labels according to plan["language"]
- If plan["language"] == "es": use "Frecuencia" and "Densidad"
- If plan["language"] == "en": use "Frequency" and "Density"

KIND -> IMPLEMENTATION:
- correlation_heatmap: use ONLY num_features, sns.heatmap(corr, annot=False, cmap='coolwarm')
- histogram: sns.histplot(df[x].dropna(), kde=True)
- boxplot: sns.boxplot(data=df, ...)
- scatter: sns.scatterplot(data=df, x=x, y=y, hue=hue if hue else None)

STATS TESTS (plan["stats"]):
- pearson: stats.pearsonr(x, y) -> print r and p
- ttest: stats.ttest_ind comparing y across two groups
- anova: stats.f_oneway over groups; compute eta2
- chi-square: stats.chi2_contingency(pd.crosstab(x, y))

STAT LINE FORMAT (examples):
- STAT: kind=histogram file=dist_age.png x=age n=1000 mean=... std=... missing=...
- STAT: kind=correlation_heatmap file=correlation.png top_pairs=[(...)]
- STAT: test=anova y=math_score group=race_ethnicity p=0.002 eta2=0.11
- STAT: test=pearson x=math_score y=reading_score r=0.62 p=1.2e-10

OUTPUT: Only Python code. NO markdown. NO backticks. NO explanations.
"""
        
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": "You are a code generation engine. Output only executable Python code."},
                {"role": "user", "content": prompt}
            ],
        )
        
        if hasattr(response, "usage") and response.usage is not None:
            self.tracker.record(self.model, response.usage)
        
        return response.choices[0].message.content.strip()

# ======================================================
# ReAct Orchestrator
# ======================================================
class ReActAnalyzer:
    def __init__(self, runner, decider, coder, tracker, max_steps=20, verbose=False, output_filename="report.md", progress_callback=None):
        self.runner = runner
        self.decider = decider
        self.coder = coder
        self.tracker = tracker
        self.max_steps = max_steps
        self.verbose = verbose
        self.output_filename = output_filename
        self.progress_callback = progress_callback
        self.observations = []
        self.analysis_log = []
        self.error_history = []
        self.run_start_fig_count = len(list(self.runner.figures_dir.glob("*.png")))
        self.coverage = {
            "histogram": 0,
            "correlation_heatmap": 0,
            "boxplot": 0,
            "scatter": 0,
            "stats": 0
        }
        self.used_cols = {
            "histogram": set(),
            "scatter": set(),
            "boxplot": set()
        }
        self.data_preview = self.runner.df.head(10).to_string()
        self.datacard = ReActAnalyzer.build_datacard(
            self.runner.df, self.runner.num_features, self.runner.cat_features
        )
    
    @staticmethod
    def build_datacard(df, num_features, cat_features):
        """Build compact datacard"""
        shape = df.shape
        missing = (df.isna().mean().sort_values(ascending=False).head(8) * 100).round(1)
        cat_card = {}
        for c in cat_features[:8]:
            try:
                cat_card[c] = int(df[c].nunique(dropna=True))
            except Exception:
                cat_card[c] = None
        
        num_limited = num_features[:25]
        cat_limited = cat_features[:25]
        
        return (
            f"shape={shape}\n"
            f"num_features={num_limited}\n"
            f"cat_features={cat_limited}\n"
            f"num_features_total={len(num_features)}\n"
            f"cat_features_total={len(cat_features)}\n"
            f"top_missing_pct={missing.to_dict()}\n"
            f"cat_nunique={cat_card}\n"
        )
    
    def observe(self):
        """Observe current state"""
        figs = sorted(self.runner.figures_dir.glob("*.png"))
        obs = f"Figures: {len(figs)} - " + ", ".join([f.name for f in figs]) if figs else "No figures"
        run_fig_count = max(0, len(figs) - self.run_start_fig_count)
        obs = f"{obs}\nRun figures: {run_fig_count}\nCoverage: {self.coverage}\nUsed columns: {self.used_cols}"
        self.observations.append(obs)
    
    def _update_progress(self, step, status, details=""):
        """Update progress via callback"""
        if self.progress_callback:
            self.progress_callback(step, self.max_steps, status, details)
    
    def run(self, user_requirements):
        """Main ReAct loop with progress tracking"""
        print(f"\n🚀 STARTING ANALYSIS\n{'='*80}")
        self._update_progress(0, "Initializing", "Setting up analysis environment")
        
        self.observe()
        consecutive_errors = 0
        max_consecutive_errors = 3
        
        for step in range(self.max_steps):
            print(f"\n📍 STEP {step + 1}/{self.max_steps}")
            self._update_progress(step + 1, "Running", f"Executing analysis step {step + 1}")
            
            error_context = ""
            if self.error_history:
                recent = self.error_history[-3:]
                error_context = "\n\nRECENT ERRORS:\n" + "\n".join([
                    f"- {e['error']}\n  From: {str(e.get('action',''))[:100]}..."
                    for e in recent
                ])
            
            # Get decision
            fig_count = len(list(self.runner.figures_dir.glob("*.png")))
            run_fig_count = max(0, fig_count - self.run_start_fig_count)
            
            if step == 0 and run_fig_count == 0 and (self.runner.num_features or self.runner.cat_features):
                base_num = self.runner.num_features
                base_cat = self.runner.cat_features
                analyses = []

                if len(base_num) >= 2:
                    x1 = base_num[0]
                    x2 = base_num[1]
                    analyses = [
                        {"kind": "correlation_heatmap", "filename": "corr_1.png",
                         "title": "Correlación", "caption": "Mapa de correlación entre variables numéricas."},
                        {"kind": "histogram", "x": x1, "filename": "hist_1.png",
                         "title": "Distribución", "caption": f"Distribución de {x1}."},
                        {"kind": "scatter", "x": x1, "y": x2,
                         "filename": "scatter_1.png", "title": "Relación entre variables",
                         "caption": "Relación entre dos variables numéricas."}
                    ]
                elif len(base_num) == 1:
                    x1 = base_num[0]
                    analyses = [
                        {"kind": "histogram", "x": x1, "filename": "hist_1.png",
                         "title": "Distribución", "caption": f"Distribución de {x1}."}
                    ]
                    if base_cat:
                        analyses.append({
                            "kind": "boxplot", "y": x1, "x": base_cat[0], "filename": "box_1.png",
                            "title": "Distribución por grupo", "caption": f"Distribución de {x1} por {base_cat[0]}."
                        })
                else:
                    if base_cat:
                        x1 = base_cat[0]
                        analyses = [
                            {"kind": "histogram", "x": x1, "filename": "hist_1.png",
                             "title": "Distribución", "caption": f"Distribución de {x1}."}
                        ]

                plan = {
                    "stop": False,
                    "language": self.decider.lang,
                    "analyses": analyses,
                    "stats": []
                }
            else:
                obs_compact = "\n".join(self.observations[-2:])
                decision = self.decider.decide(
                    observations=obs_compact + error_context,
                    step_num=step,
                    max_steps=self.max_steps,
                    num_features=self.runner.num_features,
                    cat_features=self.runner.cat_features,
                    data_preview=self.data_preview
                )
                
                try:
                    plan = json.loads(decision)
                    plan = sanitize_plan(plan, self.runner.num_features, self.runner.cat_features)
                    consecutive_errors = 0
                except Exception as e:
                    consecutive_errors += 1
                    print(f"⚠️ JSON parsing error: {e}")
                    self._update_progress(step + 1, "Error", f"JSON parsing failed: {str(e)[:50]}")
                    
                    if consecutive_errors >= max_consecutive_errors:
                        print(f"❌ Too many consecutive errors ({max_consecutive_errors}), aborting")
                        break
                    
                    self.error_history.append({
                        "step": step + 1,
                        "action": "parse_plan",
                        "error": f"Invalid JSON: {str(e)}"
                    })
                    self.observations.append("ERROR: invalid_json_plan")
                    continue
            
            # Post-sanitize: avoid reusing columns
            def _pick_alt(cols, used):
                for c in cols:
                    if c not in used:
                        return c
                return None
            
            num_cols = self.runner.num_features
            any_cols = self.runner.num_features + self.runner.cat_features
            
            for a in plan.get("analyses", []):
                kind = a.get("kind")
                if kind == "histogram":
                    x = a.get("x")
                    if x in self.used_cols["histogram"]:
                        alt = _pick_alt(num_cols or any_cols, self.used_cols["histogram"])
                        if alt:
                            a["x"] = alt
                elif kind == "scatter":
                    x, y = a.get("x"), a.get("y")
                    if x in self.used_cols["scatter"] or y in self.used_cols["scatter"]:
                        candidates = [c for c in num_cols if c not in self.used_cols["scatter"]]
                        if len(candidates) >= 2:
                            a["x"], a["y"] = candidates[0], candidates[1]
                elif kind == "boxplot":
                    y = a.get("y")
                    if y in self.used_cols["boxplot"]:
                        alt = _pick_alt(num_cols, self.used_cols["boxplot"])
                        if alt:
                            a["y"] = alt
            
            # Validate plan
            try:
                validate_plan(plan, self.runner.num_features, self.runner.cat_features)
            except Exception as e:
                consecutive_errors += 1
                print(f"⚠️ Plan validation error: {e}")
                self._update_progress(step + 1, "Error", f"Plan validation failed: {str(e)[:50]}")
                
                if consecutive_errors >= max_consecutive_errors:
                    print(f"❌ Too many consecutive validation errors, aborting")
                    break
                
                self.error_history.append({
                    "step": step + 1,
                    "action": "validate_plan",
                    "error": str(e)
                })
                self.observations.append(f"ERROR: invalid_plan: {e}")
                continue
            
            # Check if analysis should stop
            fig_count = len(list(self.runner.figures_dir.glob("*.png")))
            run_fig_count = max(0, fig_count - self.run_start_fig_count)
            
            if plan.get("stop") is True and run_fig_count < 2 and self.coverage["stats"] == 0 and step < 1:
                plan["stop"] = False
                if not plan.get("analyses"):
                    fallback_x = None
                    if self.runner.num_features:
                        fallback_x = self.runner.num_features[0]
                    elif self.runner.cat_features:
                        fallback_x = self.runner.cat_features[0]
                    
                    plan["analyses"] = [
                        {"kind": "histogram", "x": fallback_x, "filename": "auto_hist_1.png",
                         "title": "Distribution", "caption": "Auto-generated distribution plot."},
                        {"kind": "correlation_heatmap", "filename": "auto_corr_1.png",
                         "title": "Correlation heatmap", "caption": "Auto-generated correlation heatmap."}
                    ]
            
            if plan.get("stop") is True:
                print("✅ Analysis complete (stop=true)")
                self._update_progress(step + 1, "Complete", "Analysis finished successfully")
                break
            
            # Generate code
            last_error = self.error_history[-1]["error"] if self.error_history else None
            code = self.coder.generate(
                plan,
                self.runner.num_features,
                self.runner.cat_features,
                self.datacard,
                error_context=last_error
            )
            
            # Execute code
            try:
                self._update_progress(step + 1, "Executing", "Running generated analysis code")
                stdout_text = self.runner.run(code, plan=plan)
                
                expected = {a["filename"] for a in plan.get("analyses", [])}
                existing = {p.name for p in self.runner.figures_dir.glob("*.png")}
                missing = expected - existing
                
                if missing:
                    print(f"⚠️  Warning: Expected figures not generated: {sorted(missing)}")
                
                if not existing and expected:
                    raise RuntimeError("No figures generated")
                
                if self.verbose:
                    print("✅ Success")
                
                stat_lines = parse_stat_lines(stdout_text)
                test_stat_count = sum(
                    1 for ln in stat_lines
                    if "STAT:" in ln and "test=" in ln and "skipped" not in ln
                )
                
                for a in plan.get("analyses", []):
                    fn = a.get("filename")
                    kind = a.get("kind")
                    if fn and (self.runner.figures_dir / fn).exists():
                        if kind in self.coverage:
                            self.coverage[kind] += 1
                        if kind == "histogram":
                            x = a.get("x")
                            if x:
                                self.used_cols["histogram"].add(x)
                        elif kind == "scatter":
                            x, y = a.get("x"), a.get("y")
                            if x:
                                self.used_cols["scatter"].add(x)
                            if y:
                                self.used_cols["scatter"].add(y)
                        elif kind == "boxplot":
                            y = a.get("y")
                            if y:
                                self.used_cols["boxplot"].add(y)
                
                if test_stat_count:
                    self.coverage["stats"] += test_stat_count
                
                self.observe()
                consecutive_errors = 0
                
                self.analysis_log.append({
                    "step": step + 1,
                    "plan": plan,
                    "result": self.observations[-1],
                    "status": "success",
                    "stdout": stdout_text,
                    "stat_lines": stat_lines
                })
                
                fig_count = len(list(self.runner.figures_dir.glob("*.png")))
                run_fig_count = max(0, fig_count - self.run_start_fig_count)
                
                if run_fig_count >= 5 and self.coverage["stats"] >= 3:
                    print("🎯 Auto-stop: reached target coverage")
                    self._update_progress(step + 1, "Complete", "Target coverage reached")
                    break
                    
            except Exception as e:
                consecutive_errors += 1
                error_msg = str(e)
                print(f"❌ Execution error: {error_msg}")
                self._update_progress(step + 1, "Error", f"Execution failed: {error_msg[:50]}")
                
                if consecutive_errors >= max_consecutive_errors:
                    print(f"❌ Too many consecutive execution errors, aborting")
                    break
                
                self.error_history.append({
                    "step": step + 1,
                    "action": plan,
                    "error": error_msg
                })
                error_obs = f"ERROR: {error_msg}"
                self.observations.append(error_obs)
                self.analysis_log.append({
                    "step": step + 1,
                    "plan": plan,
                    "result": error_obs,
                    "status": "error"
                })
        
        # Generate final report
        self._update_progress(self.max_steps, "Generating Report", "Creating final analysis report")
        report_text = self.decider.synthesize_report(
            self.analysis_log,
            user_requirements,
            self.runner.figures_dir,
            self.data_preview,
            lang=self.decider.lang
        )
        
        report_path = self.runner.output_dir / self.output_filename
        report_path.write_text(report_text, encoding='utf-8')
        print(f"✨ Report saved: {report_path}")
        self._update_progress(self.max_steps, "Complete", f"Report saved: {report_path.name}")
        
        return report_path, report_text

# ======================================================
# Runtime helpers
# ======================================================

def run_analysis(
    input_path: Path,
    output_path: Path,
    output_format: str = None,
    api_key: str = None,
    verbose: bool = False,
    custom_prompt: str = None,
    prompt_file: Path = None,
    lang: str = "es",
    clean_figures: bool = False,
    auto_fix: bool = None,
    progress_callback=None,
):
    api_key = get_api_key(api_key)
    output_dir = output_path.parent
    output_dir.mkdir(parents=True, exist_ok=True)

    if clean_figures:
        robust_rmtree(output_dir / "figures")

    # Load custom prompt
    if prompt_file and custom_prompt is None:
        if prompt_file.exists():
            custom_prompt = prompt_file.read_text(encoding='utf-8')
            print(f"📄 Loaded custom prompt from: {prompt_file}")
        else:
            print(f"⚠️  Prompt file not found: {prompt_file}")

    # Initialize tracker
    tracker = UsageTracker()

    # Load data
    df, num_features, cat_features = load_and_clean(input_path)

    figs_before = set((output_dir / "figures").glob("*.png"))

    # Validate data quality
    validator = DataQualityValidator(df, num_features, cat_features, verbose=verbose)
    validation_passed = validator.validate()

    # Auto-fix if needed
    if not validation_passed:
        if auto_fix is None:
            print("\n⚠️  Data quality issues detected!")
            response = input("Attempt automatic fixes? (y/n): ").strip().lower()
            auto_fix = (response == 'y')

        if auto_fix:
            df = validator.auto_fix()
            print("✅ Using corrected dataset")
        else:
            print("⚠️  Proceeding with original data")

    # Create runner
    runner = SafeRunner(df, num_features, cat_features, output_dir, verbose=verbose)

    # Create agents
    decider = DecisionMaker(api_key, tracker, custom_prompt=custom_prompt, lang=lang, verbose=verbose)
    coder = CodexGenerator(api_key, tracker, verbose=verbose)

    # -----------------------------------------------------
    # TITLE IMAGE (PEXELS) - portada
    # -----------------------------------------------------
    title_path = runner.figures_dir / "title_image.jpg"
    data_preview_for_image = df.head(10).to_string()
    img_prompt = decider.generate_title_image_prompt(
        data_preview=data_preview_for_image,
        num_features=num_features,
        cat_features=cat_features,
        dataset_name=input_path.name
    )
    print("PROMPT (title image):", img_prompt)

    pexels_key = os.getenv("PEXELS_API_KEY")
    if not pexels_key:
        print("PEXELS_API_KEY not found. Skipping title image.")
    else:
        try:
            from RetrieveImage import download_image
            download_image(
                search_query=img_prompt,
                api_key=pexels_key,
                file_name=str(title_path)
            )
            print("Title image downloaded:", title_path.exists(), "-", title_path.resolve())
        except Exception as e:
            print("Could not download title image:", str(e)[:200])

    analyzer = ReActAnalyzer(
        runner, decider, coder, tracker, 
        max_steps=5, verbose=verbose, 
        output_filename=output_path.name,
        progress_callback=progress_callback
    )

    user_requirements = custom_prompt or "Comprehensive domain-specific report with embedded figures."

    # Run analysis
    report_md_path, report_text = analyzer.run(user_requirements)

    # -------------------------
    # INSERT TITLE IMAGE INTO report.md (below H1)
    # -------------------------
    if title_path.exists():
        title_image_md = (
            "\n::: {style=\"text-align:center;\"}\n"
            "![](figures/title_image.jpg){ width=80% }\n"
            ":::\n"
        )

        lines = report_text.splitlines()
        inserted = False
        for i, line in enumerate(lines):
            if line.startswith("# "):
                lines.insert(i + 1, title_image_md)
                inserted = True
                break

        if inserted:
            report_text = "\n".join(lines)
            report_md_path.write_text(report_text, encoding="utf-8")
            print("Title image inserted in report.md")
        else:
            print("No H1 (# ) found. Title image not inserted.")

    out_file = None
    if output_format:
        print(f"🔄 Converting to {output_format}...")
        out_file = output_path.with_suffix(f".{output_format}")

        if output_format == "pptx":
            # ======================================================
            # PPTX usando md_to_ppt.py (NO Pandoc)
            # ======================================================

            # Script md_to_ppt.py (misma carpeta que este archivo)
            md_to_ppt_script = Path(__file__).with_name("md_to_ppt.py")

            # Template PPTX (misma carpeta, nombre fijo por ahora)
            template_pptx = Path(__file__).with_name("template.pptx")

            if not md_to_ppt_script.exists():
                raise FileNotFoundError(
                    f"No se encontró md_to_ppt.py en {md_to_ppt_script}"
                )

            if not template_pptx.exists():
                raise FileNotFoundError(
                    f"No se encontró template.pptx en {template_pptx}"
                )

            cmd = [
                sys.executable,
                str(md_to_ppt_script),
                "--md", str(report_md_path),
                "--pptx", str(template_pptx),
                "--api_key", str(api_key),
                "--output", str(out_file.name),
            ]

            try:
                subprocess.run(
                    cmd,
                    cwd=str(output_dir),
                    check=True,
                    capture_output=True,
                    text=True
                )
                print(f"✅ Presentation saved to {out_file}")
            except subprocess.CalledProcessError as e:
                err = (e.stderr or e.stdout or str(e))[:2000]
                raise RuntimeError(f"md_to_ppt.py failed:\n{err}")

        else:
            try:
                subprocess.run(
                    ["pandoc", report_md_path.name, "-o", out_file.name],
                    cwd=output_dir,
                    check=True
                )
                print(f"✅ Document saved to {out_file}")
            except FileNotFoundError:
                print("❌ Pandoc not installed or not in PATH")
            except subprocess.CalledProcessError as e:
                print(f"❌ Pandoc failed: {e}")

    print("\n" + "="*70)
    tracker.print_summary()
    tracker.save_log(output_dir)

    figs_after = set(runner.figures_dir.glob("*.png"))
    new_figs = sorted([p for p in figs_after if p not in figs_before])

    return {
        "report_md_path": report_md_path,
        "report_text": report_text,
        "figures_dir": runner.figures_dir,
        "new_figs": new_figs,
        "out_file": out_file,
    }

# ======================================================
# CLI
# ======================================================

def main():
    parser = argparse.ArgumentParser(
        description="IAthon - Intelligent Automated Data Analysis"
    )
    parser.add_argument("-i", "--input", required=True, help="Input data file (CSV or Excel)")
    parser.add_argument("-o", "--output", required=True, help="Output file path")
    parser.add_argument("-f", "--format", choices=["pdf", "docx", "pptx"], help="Output format")
    parser.add_argument("--api-key", required=False, help="OpenAI API key")
    parser.add_argument("-v", "--verbose", action="store_true", help="Verbose output")
    parser.add_argument("--lang", choices=["es", "en"], default="es", help="Report language")
    parser.add_argument("--clean-figures", action="store_true", help="Clean figures directory")
    parser.add_argument("-p", "--prompt", required=False, help="Custom prompt")
    parser.add_argument("--prompt-file", required=False, help="Path to prompt file")
    
    args = parser.parse_args()

    custom_prompt = args.prompt
    prompt_file = Path(args.prompt_file) if args.prompt_file else None

    run_analysis(
        input_path=Path(args.input),
        output_path=Path(args.output),
        output_format=args.format,
        api_key=args.api_key,
        verbose=args.verbose,
        custom_prompt=custom_prompt,
        prompt_file=prompt_file,
        lang=args.lang,
        clean_figures=args.clean_figures,
        auto_fix=None,
    )

if __name__ == "__main__":
    main()
