import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime


plt.style.use('default')
sns.set_palette("husl")


df = pd.read_excel('ROMI_center_Данные_для_тестового_задания_аналитик_данных.xlsx', 
                   sheet_name='Сквозная аналитика 2025')

df['source'] = np.where(df['traffic_source'] == 'ad', df['utm_source'], df['traffic_source'])


df['source'] = df['source'].fillna('other')

# Агрегируем данные по источникам
agg_data = df.groupby('source').agg({
    'costs': 'sum',
    'clicks': 'sum',
    'visits': 'sum',
    'leads': 'sum',
    'q_leads': 'sum',
    'orders': 'sum',
    'revenue': 'sum'
}).reset_index()

# Рассчитываем дополнительные метрики
agg_data['total_leads'] = agg_data['leads'] + agg_data['q_leads']
agg_data['cpa'] = agg_data['costs'] / agg_data['total_leads']
agg_data['cpo'] = agg_data['costs'] / agg_data['orders']
agg_data['roas'] = agg_data['revenue'] / agg_data['costs']


agg_data.replace([np.inf, -np.inf], np.nan, inplace=True)


fig, axes = plt.subplots(2, 3, figsize=(18, 12))
fig.suptitle('Анализ эффективности рекламных источников', fontsize=16, fontweight='bold')

# 1. Распределение расходов по источникам
costs_data = agg_data[agg_data['costs'] > 0]
axes[0, 0].pie(costs_data['costs'], labels=costs_data['source'], autopct='%1.1f%%')
axes[0, 0].set_title('Распределение расходов по источникам')

# 2. ROAS по источникам
roas_data = agg_data[agg_data['roas'].notna()].sort_values('roas', ascending=False)
axes[0, 1].bar(roas_data['source'], roas_data['roas'])
axes[0, 1].set_title('ROAS по источникам')
axes[0, 1].tick_params(axis='x', rotation=45)

# 3. Сравнение расходов и доходов по источникам
width = 0.35
x = np.arange(len(roas_data))
axes[0, 2].bar(x - width/2, roas_data['costs'], width, label='Расходы')
axes[0, 2].bar(x + width/2, roas_data['revenue'], width, label='Доход')
axes[0, 2].set_title('Сравнение расходов и доходов')
axes[0, 2].set_xticks(x)
axes[0, 2].set_xticklabels(roas_data['source'], rotation=45)
axes[0, 2].legend()

# 4. Стоимость заявки (CPA) по источникам
cpa_data = agg_data[agg_data['cpa'].notna()].sort_values('cpa', ascending=False)
axes[1, 0].bar(cpa_data['source'], cpa_data['cpa'])
axes[1, 0].set_title('Стоимость заявки (CPA) по источникам')
axes[1, 0].tick_params(axis='x', rotation=45)

# 5. Количество заявок и продаж по источникам
leads_orders_data = agg_data.sort_values('total_leads', ascending=False)
x = np.arange(len(leads_orders_data))
axes[1, 1].bar(x - width/2, leads_orders_data['total_leads'], width, label='Заявки')
axes[1, 1].bar(x + width/2, leads_orders_data['orders'], width, label='Продажи')
axes[1, 1].set_title('Заявки и продажи по источникам')
axes[1, 1].set_xticks(x)
axes[1, 1].set_xticklabels(leads_orders_data['source'], rotation=45)
axes[1, 1].legend()

# 6. Эффективность каналов 
axes[1, 2].axis('off')
table_data = agg_data[['source', 'costs', 'revenue', 'roas']].sort_values('roas', ascending=False)
table = axes[1, 2].table(cellText=table_data.round(2).values,
                         colLabels=table_data.columns,
                         cellLoc='center',
                         loc='center')
table.auto_set_font_size(False)
table.set_fontsize(9)
table.scale(1, 1.5)
axes[1, 2].set_title('Общая эффективность источников')

plt.tight_layout()
plt.savefig('marketing_dashboard.png', dpi=300, bbox_inches='tight')
plt.show()


summary_table = agg_data[['source', 'costs', 'clicks', 'total_leads', 'cpa', 'orders', 'revenue', 'roas']]
summary_table.to_csv('marketing_summary.csv', index=False, float_format='%.2f')

print("Файлы marketing_dashboard.png и marketing_summary.csv созданы.")