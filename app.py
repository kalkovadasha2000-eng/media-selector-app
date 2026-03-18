import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.colors import LinearSegmentedColormap
import statsmodels.api as sm
from statsmodels.formula.api import ols
from scipy import stats
import warnings
warnings.filterwarnings('ignore')

# ============================================
# LANGUAGE SETTINGS
# ============================================

LANGUAGES = {
    'Русский': 'ru',
    'English': 'en'
}

TEXTS = {
    'ru': {
        'app_title': "🔬 Media Selector Pro - Полный статистический анализ",
        'load_data': "📂 Загрузка данных",
        'select_excel': "Выберите Excel файл",
        'loaded': "✅ Загружено",
        'rows': "строк",
        'combinations': "комбинаций",
        'Titer': 'Титр (г/л)',
        'Cost': 'Стоимость (CNY/g)',
        'Qp': 'Уд. продуктивность',
        'Acid': 'Кислые формы (%)',
        'Main': 'Основные формы (%)',
        'HM': 'Высокоманнозные (%)',
        'Fuc': 'Фукозилированные (%)',
        'Gal': 'Галактозилированные (%)',
        'Viability': 'Жизнеспособность (%)',
        'DT3': 'Время удвоения (ч)',
        'higher_better': '↑',
        'lower_better': '↓',
        'tab_data': "📋 Данные",
        'tab_stats': "📊 Статистика",
        'tab_graphs': "📈 Графики",
        'tab_selection': "🎯 Подбор",
        'tab_all_params': "🔥 Все параметры",
        'tab_integral': "⚖️ Интегральная карта",
        'tab_rating': "🏆 Рейтинг",
        'raw_data': "Исходные данные (с блоками)",
        'stat_significance': "Статистическая значимость",
        'anova_table': "Таблица ANOVA",
        'detailed_anova': "Детальный ANOVA анализ",
        'parameter': "Параметр",
        'source': "Источник",
        'sum_sq': "Сумма квадратов",
        'df': "ст.св.",
        'mean_sq': "Ср. квадрат",
        'F': "F-значение",
        'p_value': "p-value",
        'media': "Среда",
        'feed': "Подпитка",
        'interaction': "Взаимодействие",
        'block': "Блок",
        'residual': "Остаток",
        'show_correlations': "Показать корреляции",
        'correlation_matrix': "Матрица корреляций",
        'select_parameter': "Выберите параметр",
        'interaction_plot': "📈 График взаимодействия",
        'box_plot': "📦 Box plot",
        'error_bar': "📊 С ошибками",
        'by_media': "по средам",
        'by_feed': "по подпиткам",
        'best_combo': "🏆 Лучшая комбинация",
        'worst_combo': "📉 Худшая комбинация",
        'value': "значение",
        'target_profile': "🎯 Подбор по целевому профилю",
        'enter_targets': "Введите целевые значения",
        'recommendations': "Рекомендации",
        'min': "Мин",
        'max': "Макс",
        'get_recommendations': "🎯 Получить рекомендации",
        'hits': "Попаданий",
        'match_percent': "Соответствие_%",
        'all_params_values': "Все параметры",
        'all_params_heatmap': "🔥 Тепловая карта всех параметров",
        'show_unified_heatmap': "Показать единую тепловую карту",
        'green_better': "(зеленый = лучше)",
        'integral_heatmap': "⚖️ Интегральная карта с весами",
        'set_priorities': "Настройте приоритеты параметров:",
        'apply_weights': "✅ Применить веса и построить карту",
        'show_current': "👁️ Показать с текущими весами",
        'weights': "Веса",
        'direction': "Направление",
        'all_combinations': "🏆 Все комбинации с рейтингом",
        'upload_first': "👈 Загрузите файл с данными для начала работы",
        'no_params': "Нет доступных параметров для отображения",
        'stars_note': "* p<0.05 | ** p<0.01 | *** p<0.001 | ns - не значимо",
    },
    'en': {
        'app_title': "🔬 Media Selector Pro - Complete Statistical Analysis",
        'load_data': "📂 Load Data",
        'select_excel': "Choose Excel file",
        'loaded': "✅ Loaded",
        'rows': "rows",
        'combinations': "combinations",
        'Titer': 'Titer (g/L)',
        'Cost': 'Cost (CNY/g)',
        'Qp': 'Specific productivity',
        'Acid': 'Acidic forms (%)',
        'Main': 'Main forms (%)',
        'HM': 'High-mannose (%)',
        'Fuc': 'Fucosylated (%)',
        'Gal': 'Galactosylated (%)',
        'Viability': 'Viability (%)',
        'DT3': 'Doubling time (h)',
        'higher_better': '↑',
        'lower_better': '↓',
        'tab_data': "📋 Data",
        'tab_stats': "📊 Statistics",
        'tab_graphs': "📈 Graphs",
        'tab_selection': "🎯 Selection",
        'tab_all_params': "🔥 All Parameters",
        'tab_integral': "⚖️ Integral Map",
        'tab_rating': "🏆 Rating",
        'raw_data': "Raw data (with blocks)",
        'stat_significance': "Statistical Significance",
        'anova_table': "ANOVA Table",
        'detailed_anova': "Detailed ANOVA Analysis",
        'parameter': "Parameter",
        'source': "Source",
        'sum_sq': "Sum Sq",
        'df': "df",
        'mean_sq': "Mean Sq",
        'F': "F-value",
        'p_value': "p-value",
        'media': "Media",
        'feed': "Feed",
        'interaction': "Interaction",
        'block': "Block",
        'residual': "Residual",
        'show_correlations': "Show correlations",
        'correlation_matrix': "Correlation Matrix",
        'select_parameter': "Select parameter",
        'interaction_plot': "📈 Interaction plot",
        'box_plot': "📦 Box plot",
        'error_bar': "📊 Error bars",
        'by_media': "by media",
        'by_feed': "by feed",
        'best_combo': "🏆 Best combination",
        'worst_combo': "📉 Worst combination",
        'value': "value",
        'target_profile': "🎯 Target Profile Selection",
        'enter_targets': "Enter target values",
        'recommendations': "Recommendations",
        'min': "Min",
        'max': "Max",
        'get_recommendations': "🎯 Get recommendations",
        'hits': "Hits",
        'match_percent': "Match_%",
        'all_params_values': "All parameters",
        'all_params_heatmap': "🔥 Heatmap of all parameters",
        'show_unified_heatmap': "Show unified heatmap",
        'green_better': "(green = better)",
        'integral_heatmap': "⚖️ Integral Heatmap with weights",
        'set_priorities': "Set parameter priorities:",
        'apply_weights': "✅ Apply weights and show map",
        'show_current': "👁️ Show with current weights",
        'weights': "Weights",
        'direction': "Direction",
        'all_combinations': "🏆 All combinations with rating",
        'upload_first': "👈 Upload a data file to start",
        'no_params': "No parameters available to display",
        'stars_note': "* p<0.05 | ** p<0.01 | *** p<0.001 | ns - not significant",
    }
}

# ============================================
# PARAMETER INFO
# ============================================

PARAM_INFO = {
    'Titer': {'default_weight': 5, 'direction': '↑'},
    'Cost': {'default_weight': 4, 'direction': '↓'},
    'Qp': {'default_weight': 4, 'direction': '↑'},
    'Acid': {'default_weight': 3, 'direction': '↓'},
    'Main': {'default_weight': 3, 'direction': '↑'},
    'HM': {'default_weight': 2, 'direction': '↓'},
    'Fuc': {'default_weight': 2, 'direction': '↓'},
    'Gal': {'default_weight': 1, 'direction': '↑'},
    'Viability': {'default_weight': 1, 'direction': '↑'},
    'DT3': {'default_weight': 1, 'direction': '↓'}
}

COLUMN_MAPPING = {
    'Уд. Прод.': 'Qp',
    'Стоимость': 'Cost',
    'Титр': 'Titer',
    'Относительное содержание кислой фракции': 'Acid',
    'Относительное содержание основной фракции': 'Main',
    'Относительное содержание высокоманнозных гликанов': 'HM',
    'Относительное содержание фукозилированных гликанов': 'Fuc',
    'Относительное содержание галактозилированных гликанов': 'Gal',
    'Жизнеспособность на последний день': 'Viability',
    'Время удвоения на 3 день': 'DT3'
}

ENV_ORDER = ['Altair', 'Star', 'B601', 'B100']
FEED_ORDER = ['AF183', 'AF169', 'F602as', 'F100as']

FEED_COLORS = {
    'AF183': '#1f77b4',
    'AF169': '#ff7f0e',
    'F602as': '#2ca02c',
    'F100as': '#d62728'
}

# ============================================
# MAIN CORE CLASS
# ============================================

class MediaSelectorCore:
    def __init__(self):
        self.df = None
        self.df_mean = None
        self.df_norm = None
        self.significance = {}
        self.anova_detailed = {}
        self.correlations = None
        self.current_weights = {param: info['default_weight'] for param, info in PARAM_INFO.items()}
        self.block_col = None
    
    def load_data(self, uploaded_file):
        try:
            self.df = pd.read_excel(uploaded_file)
            
            # Rename columns
            rename_dict = {}
            for old, new in COLUMN_MAPPING.items():
                if old in self.df.columns:
                    rename_dict[old] = new
            self.df.rename(columns=rename_dict, inplace=True)
            
            # Determine media and feed columns
            env_col = 'Среда' if 'Среда' in self.df.columns else 'срекда'
            feed_col = 'Подпитка' if 'Подпитка' in self.df.columns else 'подпитка'
            
            if 'Блок' in self.df.columns:
                self.block_col = 'Блок'
            else:
                self.block_col = None
            
            self.df['Комбинация'] = self.df[env_col] + ' + ' + self.df[feed_col]
            
            # Average for tables
            numeric_cols = self.df.select_dtypes(include=[np.number]).columns
            self.df_mean = self.df.groupby([env_col, feed_col])[numeric_cols].mean().reset_index()
            self.df_mean.rename(columns={env_col: 'Среда', feed_col: 'Подпитка'}, inplace=True)
            self.df_mean['Комбинация'] = self.df_mean['Среда'] + ' + ' + self.df_mean['Подпитка']
            
            self.calculate_significance()
            self.calculate_detailed_anova()
            self.calculate_correlations()
            self.normalize_data()
            
            return True
        except Exception as e:
            st.error(f"Error loading data: {e}")
            return False
    
    def calculate_significance(self):
        """Calculate p-values for main factors"""
        self.significance = {}
        for param in PARAM_INFO.keys():
            if param in self.df.columns:
                try:
                    if self.block_col:
                        formula = f'{param} ~ C(Среда) + C(Подпитка) + C({self.block_col}) + C(Среда):C(Подпитка)'
                    else:
                        formula = f'{param} ~ C(Среда) + C(Подпитка) + C(Среда):C(Подпитка)'
                    
                    model = ols(formula, data=self.df).fit()
                    anova_table = sm.stats.anova_lm(model, typ=2)
                    
                    self.significance[param] = {
                        'env': anova_table['PR(>F)']['C(Среда)'],
                        'feed': anova_table['PR(>F)']['C(Подпитка)'],
                        'inter': anova_table['PR(>F)']['C(Среда):C(Подпитка)']
                    }
                    
                    if self.block_col and f'C({self.block_col})' in anova_table.index:
                        self.significance[param]['block'] = anova_table['PR(>F)'][f'C({self.block_col})']
                except Exception as e:
                    print(f"Error calculating significance for {param}: {e}")
                    self.significance[param] = {'env': 1, 'feed': 1, 'inter': 1}
    
    def calculate_detailed_anova(self):
        """Calculate detailed ANOVA table with sum of squares, df, mean squares, F, p-value"""
        self.anova_detailed = {}
        
        for param in PARAM_INFO.keys():
            if param in self.df.columns:
                try:
                    if self.block_col:
                        formula = f'{param} ~ C(Среда) + C(Подпитка) + C({self.block_col}) + C(Среда):C(Подпитка)'
                    else:
                        formula = f'{param} ~ C(Среда) + C(Подпитка) + C(Среда):C(Подпитка)'
                    
                    model = ols(formula, data=self.df).fit()
                    anova_table = sm.stats.anova_lm(model, typ=2)
                    
                    # Print for debugging
                    print(f"ANOVA table for {param}:")
                    print(anova_table)
                    
                    detailed = []
                    
                    # Check if we have the expected columns
                    has_sum_sq = 'sum_sq' in anova_table.columns
                    has_df = 'df' in anova_table.columns
                    has_mean_sq = 'mean_sq' in anova_table.columns
                    has_F = 'F' in anova_table.columns
                    has_PR = 'PR(>F)' in anova_table.columns
                    
                    # Media
                    if 'C(Среда)' in anova_table.index:
                        row = anova_table.loc['C(Среда)']
                        detailed.append({
                            'Source': 'Media',
                            'sum_sq': float(row['sum_sq']) if has_sum_sq else 0,
                            'df': int(row['df']) if has_df else 0,
                            'mean_sq': float(row['mean_sq']) if has_mean_sq else (float(row['sum_sq']/row['df']) if has_sum_sq and has_df and row['df'] > 0 else 0),
                            'F': float(row['F']) if has_F else 0,
                            'p_value': float(row['PR(>F)']) if has_PR else 1.0
                        })
                    
                    # Feed
                    if 'C(Подпитка)' in anova_table.index:
                        row = anova_table.loc['C(Подпитка)']
                        detailed.append({
                            'Source': 'Feed',
                            'sum_sq': float(row['sum_sq']) if has_sum_sq else 0,
                            'df': int(row['df']) if has_df else 0,
                            'mean_sq': float(row['mean_sq']) if has_mean_sq else (float(row['sum_sq']/row['df']) if has_sum_sq and has_df and row['df'] > 0 else 0),
                            'F': float(row['F']) if has_F else 0,
                            'p_value': float(row['PR(>F)']) if has_PR else 1.0
                        })
                    
                    # Interaction
                    if 'C(Среда):C(Подпитка)' in anova_table.index:
                        row = anova_table.loc['C(Среда):C(Подпитка)']
                        detailed.append({
                            'Source': 'Interaction',
                            'sum_sq': float(row['sum_sq']) if has_sum_sq else 0,
                            'df': int(row['df']) if has_df else 0,
                            'mean_sq': float(row['mean_sq']) if has_mean_sq else (float(row['sum_sq']/row['df']) if has_sum_sq and has_df and row['df'] > 0 else 0),
                            'F': float(row['F']) if has_F else 0,
                            'p_value': float(row['PR(>F)']) if has_PR else 1.0
                        })
                    
                    # Block (if exists)
                    if self.block_col and f'C({self.block_col})' in anova_table.index:
                        row = anova_table.loc[f'C({self.block_col})']
                        detailed.append({
                            'Source': 'Block',
                            'sum_sq': float(row['sum_sq']) if has_sum_sq else 0,
                            'df': int(row['df']) if has_df else 0,
                            'mean_sq': float(row['mean_sq']) if has_mean_sq else (float(row['sum_sq']/row['df']) if has_sum_sq and has_df and row['df'] > 0 else 0),
                            'F': float(row['F']) if has_F else 0,
                            'p_value': float(row['PR(>F)']) if has_PR else 1.0
                        })
                    
                    # Residual
                    if 'Residual' in anova_table.index:
                        row = anova_table.loc['Residual']
                        detailed.append({
                            'Source': 'Residual',
                            'sum_sq': float(row['sum_sq']) if has_sum_sq else 0,
                            'df': int(row['df']) if has_df else 0,
                            'mean_sq': float(row['mean_sq']) if has_mean_sq else (float(row['sum_sq']/row['df']) if has_sum_sq and has_df and row['df'] > 0 else 0),
                            'F': np.nan,
                            'p_value': np.nan
                        })
                    
                    # Create DataFrame
                    if detailed:
                        self.anova_detailed[param] = pd.DataFrame(detailed)
                        print(f"Created ANOVA table for {param} with {len(detailed)} rows")
                    else:
                        # Create empty dataframe with correct structure
                        self.anova_detailed[param] = pd.DataFrame(columns=['Source', 'sum_sq', 'df', 'mean_sq', 'F', 'p_value'])
                        print(f"Empty ANOVA table for {param}")
                    
                except Exception as e:
                    print(f"Error calculating detailed ANOVA for {param}: {e}")
                    # Create empty dataframe with correct structure
                    self.anova_detailed[param] = pd.DataFrame(columns=['Source', 'sum_sq', 'df', 'mean_sq', 'F', 'p_value'])
    
    def calculate_correlations(self):
        numeric_cols = [p for p in PARAM_INFO.keys() if p in self.df.columns]
        if len(numeric_cols) > 1:
            self.correlations = self.df[numeric_cols].corr()
    
    def get_stars(self, p):
        if pd.isna(p):
            return 'ns'
        if p < 0.001: return '***'
        elif p < 0.01: return '**'
        elif p < 0.05: return '*'
        return 'ns'
    
    def update_weights(self, new_weights):
        self.current_weights = new_weights
        self.normalize_data()
    
    def normalize_data(self):
        if self.df_mean is None:
            return
        
        self.df_norm = self.df_mean.copy()
        
        # Normalize each parameter
        for param, info in PARAM_INFO.items():
            if param in self.df_norm.columns:
                values = self.df_norm[param].values
                min_val = np.min(values)
                max_val = np.max(values)
                
                if max_val > min_val:
                    if info['direction'] == '↑':
                        self.df_norm[f'{param}_norm'] = (values - min_val) / (max_val - min_val)
                    else:
                        self.df_norm[f'{param}_norm'] = (max_val - values) / (max_val - min_val)
                else:
                    self.df_norm[f'{param}_norm'] = 0.5
        
        # Calculate weighted rating with ALL parameters
        self.df_norm['Рейтинг'] = 0
        total_weight = 0
        
        for param, info in PARAM_INFO.items():
            norm_col = f'{param}_norm'
            if norm_col in self.df_norm.columns:
                weight = self.current_weights.get(param, info['default_weight'])
                self.df_norm['Рейтинг'] += self.df_norm[norm_col] * weight
                total_weight += weight
        
        if total_weight > 0:
            self.df_norm['Рейтинг'] = self.df_norm['Рейтинг'] / total_weight
            self.df_norm['Рейтинг'] = self.df_norm['Рейтинг'].round(3)
    
    def get_best_and_worst_combinations(self, param):
        if param not in self.df_mean.columns:
            return None, None
        
        if PARAM_INFO[param]['direction'] == '↑':
            best_idx = self.df_mean[param].idxmax()
            worst_idx = self.df_mean[param].idxmin()
        else:
            best_idx = self.df_mean[param].idxmin()
            worst_idx = self.df_mean[param].idxmax()
        
        best = self.df_mean.loc[best_idx]
        worst = self.df_mean.loc[worst_idx]
        
        return best, worst
    
    def plot_interaction(self, param, lang):
        if param not in self.df_mean.columns:
            return False
        
        t = TEXTS[lang]
        
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))
        
        for feed in FEED_ORDER:
            feed_data = self.df_mean[self.df_mean['Подпитка'] == feed].sort_values('Среда')
            if not feed_data.empty:
                ax1.plot(feed_data['Среда'], feed_data[param], 
                        'o-', linewidth=2, markersize=8, 
                        color=FEED_COLORS.get(feed, 'gray'), label=feed)
        
        ax1.set_xlabel('Media', fontsize=12)
        ax1.set_ylabel(t[param], fontsize=12)
        ax1.set_title(f'{t[param]} - interaction', fontsize=14)
        ax1.legend(title='Feed', bbox_to_anchor=(1.05, 1))
        ax1.grid(True, alpha=0.3)
        
        ax2.axis('off')
        text = ""
        
        if param in self.significance:
            sig = self.significance[param]
            
            text += f"📊 Statistical significance\n"
            text += f"{'='*30}\n\n"
            text += f"Media: p = {sig['env']:.4f} {self.get_stars(sig['env'])}\n"
            text += f"Feed: p = {sig['feed']:.4f} {self.get_stars(sig['feed'])}\n"
            text += f"Interaction: p = {sig['inter']:.4f} {self.get_stars(sig['inter'])}\n"
            
            if 'block' in sig:
                text += f"Block: p = {sig['block']:.4f} {self.get_stars(sig['block'])}\n"
            
            text += f"\n{'='*30}\n"
        
        best, worst = self.get_best_and_worst_combinations(param)
        if best is not None and worst is not None:
            text += f"\n{t['best_combo']}:\n"
            text += f"  {best['Комбинация']}\n"
            text += f"  {t['value']}: {best[param]:.3f}\n\n"
            
            text += f"{t['worst_combo']}:\n"
            text += f"  {worst['Комбинация']}\n"
            text += f"  {t['value']}: {worst[param]:.3f}\n\n"
        
        text += f"\n{'='*30}\n"
        text += t['stars_note']
        
        ax2.text(0.1, 0.5, text, transform=ax2.transAxes, fontsize=12,
                verticalalignment='center',
                bbox=dict(boxstyle='round', facecolor='#E3F2FD', alpha=0.8))
        
        plt.tight_layout()
        st.pyplot(fig)
        plt.close()
        return True
    
    def plot_unified_heatmap(self, lang):
        t = TEXTS[lang]
        
        available_params = [p for p in PARAM_INFO.keys() if p in self.df_mean.columns]
        
        if not available_params:
            st.warning(t['no_params'])
            return False
        
        plot_data = self.df_mean.set_index('Комбинация')[available_params].T
        
        if 'Titer' in available_params:
            titer_order = self.df_mean.sort_values('Titer', ascending=False)['Комбинация'].values
            plot_data = plot_data[titer_order]
        
        plot_data_norm = plot_data.copy()
        for param in plot_data_norm.index:
            min_val = plot_data_norm.loc[param].min()
            max_val = plot_data_norm.loc[param].max()
            if max_val > min_val:
                if PARAM_INFO[param]['direction'] == '↑':
                    plot_data_norm.loc[param] = (plot_data_norm.loc[param] - min_val) / (max_val - min_val)
                else:
                    plot_data_norm.loc[param] = 1 - (plot_data_norm.loc[param] - min_val) / (max_val - min_val)
            else:
                plot_data_norm.loc[param] = 0.5
        
        fig, ax = plt.subplots(figsize=(14, 8))
        
        sns.heatmap(plot_data_norm, annot=plot_data, fmt='.2f', 
                   cmap='RdYlGn', linewidths=1, linecolor='white',
                   cbar_kws={'label': 'Normalized value (1 = best)'},
                   ax=ax)
        
        ax.set_title(f'Heatmap of all parameters {t["green_better"]}', fontsize=16, fontweight='bold')
        ax.set_xlabel('Media + Feed combination', fontsize=12)
        ax.set_ylabel('Parameter', fontsize=12)
        plt.xticks(rotation=45, ha='right')
        
        plt.tight_layout()
        st.pyplot(fig)
        plt.close()
        return True
    
    def plot_integral_heatmap_with_weights(self, lang):
        t = TEXTS[lang]
        
        self.normalize_data()
        
        pivot = self.df_norm.pivot_table(
            values='Рейтинг',
            index='Подпитка',
            columns='Среда',
            aggfunc='mean'
        )
        
        available_envs = [e for e in ENV_ORDER if e in pivot.columns]
        available_feeds = [f for f in FEED_ORDER if f in pivot.index]
        pivot = pivot.reindex(index=available_feeds, columns=available_envs)
        
        colors = ['#d73027', '#f46d43', '#fdae61', '#fee08b', '#ffffbf', 
                  '#d9ef8b', '#a6d96a', '#66bd63', '#1a9850']
        cmap = LinearSegmentedColormap.from_list('custom', colors, N=100)
        
        fig, ax = plt.subplots(figsize=(12, 8))
        
        sns.heatmap(pivot, annot=True, fmt='.3f', cmap=cmap,
                   linewidths=1, linecolor='white', vmin=0, vmax=1,
                   cbar_kws={'label': 'Integral rating (0-1)'},
                   ax=ax)
        
        # Create title with ALL weights
        weights_list = []
        for param in PARAM_INFO.keys():
            weight = self.current_weights.get(param, PARAM_INFO[param]['default_weight'])
            weights_list.append(f"{param}:{weight}")
        
        # Split into two lines if too long
        if len(weights_list) > 5:
            weights_text1 = " | ".join(weights_list[:5])
            weights_text2 = " | ".join(weights_list[5:])
            title = f'Integral Heatmap - All Parameters\nWeights: {weights_text1}\n{weights_text2}'
        else:
            title = f'Integral Heatmap - All Parameters\nWeights: {" | ".join(weights_list)}'
        
        plt.title(title, fontsize=14, fontweight='bold')
        ax.set_xlabel('Media', fontsize=12)
        ax.set_ylabel('Feed', fontsize=12)
        
        plt.tight_layout()
        st.pyplot(fig)
        plt.close()
        return True
    
    def plot_boxplots(self, param, lang):
        if param not in self.df.columns:
            return False
        
        t = TEXTS[lang]
        
        fig, axes = plt.subplots(1, 2, figsize=(14, 5))
        self.df.boxplot(column=param, by='Среда', ax=axes[0])
        axes[0].set_title(f'{t[param]} {t["by_media"]}')
        self.df.boxplot(column=param, by='Подпитка', ax=axes[1])
        axes[1].set_title(f'{t[param]} {t["by_feed"]}')
        plt.suptitle('')
        st.pyplot(fig)
        plt.close()
        return True
    
    def plot_bar_with_error(self, param, lang):
        if param not in self.df.columns:
            return False
        
        t = TEXTS[lang]
        
        fig, axes = plt.subplots(1, 2, figsize=(14, 5))
        
        env_stats = self.df.groupby('Среда')[param].agg(['mean', 'sem']).reset_index()
        axes[0].bar(env_stats['Среда'], env_stats['mean'], yerr=env_stats['sem'], capsize=5)
        axes[0].set_title(f'{t[param]} {t["by_media"]}')
        
        feed_stats = self.df.groupby('Подпитка')[param].agg(['mean', 'sem']).reset_index()
        axes[1].bar(feed_stats['Подпитка'], feed_stats['mean'], yerr=feed_stats['sem'], capsize=5)
        axes[1].set_title(f'{t[param]} {t["by_feed"]}')
        
        st.pyplot(fig)
        plt.close()
        return True
    
    def plot_correlation_matrix(self, lang):
        if self.correlations is None:
            return False
        
        t = TEXTS[lang]
        
        fig, ax = plt.subplots(figsize=(10, 8))
        mask = np.triu(np.ones_like(self.correlations, dtype=bool))
        sns.heatmap(self.correlations, mask=mask, annot=True, fmt='.2f',
                   cmap='RdBu_r', center=0, square=True, linewidths=1)
        plt.title(t['correlation_matrix'])
        st.pyplot(fig)
        plt.close()
        return True
    
    def match_target(self, targets):
        df_match = self.df_norm.copy()
        df_match['Попаданий'] = 0
        df_match['Отклонение'] = 0.0
        df_match['Соответствие_%'] = 0.0
        
        total_params = len(targets)
        
        for param, (min_val, max_val) in targets.items():
            if param in df_match.columns:
                in_range = (df_match[param] >= min_val) & (df_match[param] <= max_val)
                df_match['Попаданий'] += in_range.astype(int)
                
                if min_val is not None and max_val is not None:
                    dev_low = np.maximum(0, (min_val - df_match[param]) / (min_val + 1e-6))
                    dev_high = np.maximum(0, (df_match[param] - max_val) / (max_val + 1e-6))
                    deviation = dev_low + dev_high
                elif min_val is not None:
                    deviation = np.maximum(0, (min_val - df_match[param]) / (min_val + 1e-6))
                elif max_val is not None:
                    deviation = np.maximum(0, (df_match[param] - max_val) / (max_val + 1e-6))
                else:
                    deviation = 0
                
                df_match['Отклонение'] += deviation
        
        if total_params > 0:
            df_match['Соответствие_%'] = (df_match['Попаданий'] / total_params * 100).round(1)
        
        return df_match.sort_values(['Попаданий', 'Отклонение', 'Рейтинг'],
                                   ascending=[False, True, False])

# ============================================
# INITIALIZATION
# ============================================

if 'core' not in st.session_state:
    st.session_state.core = MediaSelectorCore()

core = st.session_state.core

# Language selector in sidebar
with st.sidebar:
    lang = st.radio("Language / Язык", ['Русский', 'English'])
    st.session_state.lang = 'ru' if lang == 'Русский' else 'en'
    
    t = TEXTS[st.session_state.lang]
    
    st.header(t['load_data'])
    uploaded_file = st.file_uploader(t['select_excel'], type=['xlsx'])
    
    if uploaded_file is not None:
        if core.load_data(uploaded_file):
            st.success(f"{t['loaded']} {len(core.df)} {t['rows']}, {len(core.df_mean)} {t['combinations']}")

if core.df is None:
    st.info(t['upload_first'])
else:
    t = TEXTS[st.session_state.lang]
    
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        t['tab_data'],
        t['tab_stats'],
        t['tab_graphs'],
        t['tab_selection'],
        t['tab_all_params'],
        t['tab_integral'],
        t['tab_rating']
    ])
    
    with tab1:
        st.header(t['raw_data'])
        st.dataframe(core.df)
    
    with tab2:
        st.header(t['stat_significance'])
        
        sig_data = []
        for param, info in PARAM_INFO.items():
            if param in core.significance:
                sig = core.significance[param]
                sig_data.append({
                    t['parameter']: t[param],
                    t['media']: f"{sig['env']:.4f} {core.get_stars(sig['env'])}",
                    t['feed']: f"{sig['feed']:.4f} {core.get_stars(sig['feed'])}",
                    t['interaction']: f"{sig['inter']:.4f} {core.get_stars(sig['inter'])}"
                })
        if sig_data:
            st.dataframe(pd.DataFrame(sig_data))
            st.caption(t['stars_note'])
        
        st.subheader(t['detailed_anova'])
        selected_param_anova = st.selectbox(
            t['select_parameter'],
            options=[p for p in PARAM_INFO.keys() if p in core.df.columns],
            format_func=lambda x: t[x],
            key='anova_param'
        )
        
        if selected_param_anova in core.anova_detailed:
            anova_df = core.anova_detailed[selected_param_anova].copy()
            
            if not anova_df.empty:
                # Create a display copy
                anova_display = anova_df.copy()
                
                # Format numeric columns
                for col in ['sum_sq', 'mean_sq', 'F']:
                    if col in anova_display.columns:
                        anova_display[col] = pd.to_numeric(anova_display[col], errors='coerce').round(4)
                
                # Format p-values
                if 'p_value' in anova_display.columns:
                    anova_display['p_value'] = anova_display['p_value'].apply(
                        lambda x: f"{float(x):.4f} {core.get_stars(x)}" if pd.notna(x) and isinstance(x, (int, float)) else "-"
                    )
                
                # Rename columns for display
                anova_display.columns = [
                    t['source'], t['sum_sq'], t['df'], 
                    t['mean_sq'], 'F', t['p_value']
                ]
                
                st.dataframe(anova_display)
            else:
                st.info(f"No ANOVA data available for {t[selected_param_anova]}")
        else:
            st.info(f"Select a parameter to view ANOVA table")
        
        st.subheader(f"🔄 {t['correlation_matrix']}")
        if st.button(t['show_correlations']):
            core.plot_correlation_matrix(st.session_state.lang)
    
    with tab3:
        st.header(t['tab_graphs'])
        selected_param = st.selectbox(
            t['select_parameter'],
            options=[p for p in PARAM_INFO.keys() if p in core.df_mean.columns],
            format_func=lambda x: t[x]
        )
        
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button(t['interaction_plot'], use_container_width=True):
                core.plot_interaction(selected_param, st.session_state.lang)
        with col2:
            if st.button(t['box_plot'], use_container_width=True):
                core.plot_boxplots(selected_param, st.session_state.lang)
        with col3:
            if st.button(t['error_bar'], use_container_width=True):
                core.plot_bar_with_error(selected_param, st.session_state.lang)
    
    with tab4:
        st.header(t['target_profile'])
        col1, col2 = st.columns([1, 1])
        with col1:
            st.subheader(t['enter_targets'])
            targets = {}
            for param, info in PARAM_INFO.items():
                if param in core.df_norm.columns:
                    with st.expander(f"{t[param]}", expanded=False):
                        col_min, col_max = st.columns(2)
                        with col_min:
                            min_val = st.number_input(f"{t['min']}", key=f"target_{param}_min", value=0.0)
                        with col_max:
                            max_val = st.number_input(f"{t['max']}", key=f"target_{param}_max", value=100.0)
                        targets[param] = (min_val, max_val)
        with col2:
            st.subheader(t['recommendations'])
            if st.button(t['get_recommendations'], type="primary", use_container_width=True):
                if targets:
                    results = core.match_target(targets)
                    
                    display_cols = ['Среда', 'Подпитка', 'Комбинация', 'Рейтинг', 'Попаданий', 'Соответствие_%']
                    param_cols = [p for p in PARAM_INFO.keys() if p in core.df_norm.columns]
                    display_cols.extend(param_cols)
                    
                    results_display = results[display_cols].head(10).copy()
                    
                    column_names = ['Media', 'Feed', 'Combination', 'Rating', t['hits'], t['match_percent']]
                    for param in param_cols:
                        column_names.append(t[param])
                    
                    results_display.columns = column_names
                    st.dataframe(results_display, use_container_width=True)
    
    with tab5:
        st.header(t['all_params_heatmap'])
        if st.button(t['show_unified_heatmap'], use_container_width=True):
            core.plot_unified_heatmap(st.session_state.lang)
    
    with tab6:
        st.header(t['integral_heatmap'])
        
        st.subheader(t['set_priorities'])
        
        # Parameter info table
        param_info_df = pd.DataFrame([
            {
                'Parameter': t[param],
                'Direction': '↑' if info['direction'] == '↑' else '↓',
                'Goal': 'Maximize' if info['direction'] == '↑' else 'Minimize',
                'Default Weight': info['default_weight'],
                'Current Weight': core.current_weights.get(param, info['default_weight'])
            }
            for param, info in PARAM_INFO.items()
        ])
        
        st.dataframe(param_info_df, use_container_width=True)
        st.caption("↑ = чем больше, тем лучше | ↓ = чем меньше, тем лучше")
        
        st.write("### Настройте веса для всех параметров:")
        
        col1, col2 = st.columns(2)
        new_weights = {}
        
        with col1:
            st.write("**Основные параметры**")
            new_weights['Titer'] = st.slider(f"{t['Titer']} ↑", 0, 10, core.current_weights.get('Titer', 5), key='w_titer')
            new_weights['Cost'] = st.slider(f"{t['Cost']} ↓", 0, 10, core.current_weights.get('Cost', 4), key='w_cost')
            new_weights['Qp'] = st.slider(f"{t['Qp']} ↑", 0, 10, core.current_weights.get('Qp', 4), key='w_qp')
            new_weights['Acid'] = st.slider(f"{t['Acid']} ↓", 0, 10, core.current_weights.get('Acid', 3), key='w_acid')
            new_weights['Main'] = st.slider(f"{t['Main']} ↑", 0, 10, core.current_weights.get('Main', 3), key='w_main')
        
        with col2:
            st.write("**Дополнительные параметры**")
            new_weights['HM'] = st.slider(f"{t['HM']} ↓", 0, 10, core.current_weights.get('HM', 2), key='w_hm')
            new_weights['Fuc'] = st.slider(f"{t['Fuc']} ↓", 0, 10, core.current_weights.get('Fuc', 2), key='w_fuc')
            new_weights['Gal'] = st.slider(f"{t['Gal']} ↑", 0, 10, core.current_weights.get('Gal', 1), key='w_gal')
            new_weights['Viability'] = st.slider(f"{t['Viability']} ↑", 0, 10, core.current_weights.get('Viability', 1), key='w_viability')
            new_weights['DT3'] = st.slider(f"{t['DT3']} ↓", 0, 10, core.current_weights.get('DT3', 1), key='w_dt3')
        
        total_weight = sum(new_weights.values())
        st.info(f"**Сумма весов: {total_weight}**")
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button(t['apply_weights'], use_container_width=True):
                core.update_weights(new_weights)
                st.success(f"✅ Веса обновлены! Сумма весов: {total_weight}")
                core.normalize_data()
                core.plot_integral_heatmap_with_weights(st.session_state.lang)
        with col2:
            if st.button(t['show_current'], use_container_width=True):
                core.plot_integral_heatmap_with_weights(st.session_state.lang)
        
        if st.checkbox("Показать все текущие веса"):
            weights_df = pd.DataFrame([
                {'Parameter': t[param], 'Weight': core.current_weights.get(param, info['default_weight']), 'Direction': '↑' if info['direction'] == '↑' else '↓'}
                for param, info in PARAM_INFO.items()
            ])
            st.dataframe(weights_df, use_container_width=True)
    
    with tab7:
        st.header(t['all_combinations'])
        df_display = core.df_norm[['Среда', 'Подпитка', 'Комбинация', 'Рейтинг'] + 
                                   [p for p in PARAM_INFO.keys() if p in core.df_norm.columns]].copy()
        df_display = df_display.sort_values('Рейтинг', ascending=False)
        
        column_names = ['Media', 'Feed', 'Combination', 'Rating']
        for param in PARAM_INFO.keys():
            if param in core.df_norm.columns:
                column_names.append(t[param])
        
        df_display.columns = column_names
        st.dataframe(df_display)
