"""
Geração de gráficos de tendência mensal/YTD e de imagens de emoji (legendas)
usadas nos PDFs de desempenho.
"""
import gc
import os
import sys
import traceback

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import pandas as pd
from matplotlib import font_manager
from matplotlib.patches import Rectangle

from utils import clean_filename

# Tenta carregar uma fonte de emoji adequada à plataforma.
EMOJI_FONT = "Segoe UI Emoji"
if sys.platform != "win32":
    try:
        font_manager.findfont("Noto Color Emoji", fallback_to_default=False)
        EMOJI_FONT = "Noto Color Emoji"
    except Exception:
        pass


def gerar_imagem_emoji(emoji_char, bg_color, output_path, size=30, pad_inches_val=0.1):
    """
    Gera uma imagem PNG de um emoji com um fundo colorido, usada nas legendas dos PDFs.
    """
    try:
        fig, ax = plt.subplots(figsize=(0.5, 0.5), dpi=100)
        ax.add_patch(Rectangle((0, 0), 1, 1, facecolor=bg_color, edgecolor="none", transform=ax.transAxes))
        ax.text(0.5, 0.5, emoji_char, fontsize=size, ha='center', va='center',
                fontproperties=font_manager.FontProperties(family=EMOJI_FONT))
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis('off')
        plt.tight_layout(pad=0)
        plt.savefig(output_path, dpi=100, transparent=False, bbox_inches='tight', pad_inches=pad_inches_val)
        plt.close(fig)
        return True
    except Exception as e:
        print(f"Erro ao gerar imagem para emoji '{emoji_char}' para legenda: {e}")
        return False


def gerar_grafico_mensal_e_ytd(nome_cliente_para_grafico, codigo_sap_cliente, df_metrics_calculated,
                                colunas_metricas_config, pasta_graficos, log_func, metas_para_graficos_mensais,
                                ano_atual_referencia, ano_anterior_referencia, emoji_dir=None):
    """
    Gera gráficos de barras mostrando o desempenho mensal e YTD para métricas
    específicas de um cliente, incluindo comparação com o ano anterior e
    linha de meta, se aplicável.
    """
    caminhos_graficos_gerados = []

    metricas_props_map = {item['nome_original']: item for item in colunas_metricas_config}
    colunas_metricas_originais = [item['nome_original'] for item in colunas_metricas_config]

    df_cliente_metrics = df_metrics_calculated[df_metrics_calculated['RBSNO'] == codigo_sap_cliente].copy()
    df_cliente_metrics['Value'] = pd.to_numeric(df_cliente_metrics['Value'], errors='coerce').fillna(0.0)

    for metric_original in colunas_metricas_originais:
        if metric_original == "Quantity Supplied":
            continue  # Não vai para gráfico, apenas para a tabela.

        metric_props = metricas_props_map.get(metric_original, {})
        metric_display_name = metric_props.get('nome_exibicao', metric_original)
        tem_meta = metric_props.get('tem_meta', False)
        tem_emoji = metric_props.get('tem_emoji', False)

        valor_meta_para_grafico = metas_para_graficos_mensais.get(metric_original, np.nan)

        df_metric_filtered = df_cliente_metrics[df_cliente_metrics['Metric'] == metric_original].copy()

        plot_data_list = []
        ytd_current_year_val = df_metric_filtered[df_metric_filtered['Period'] == str(ano_atual_referencia)]['Value'].sum()
        ytd_previous_year_val = df_metric_filtered[df_metric_filtered['Period'] == str(ano_anterior_referencia)]['Value'].sum()

        plot_data_list.append({'Category': str(ano_anterior_referencia), 'Value': ytd_previous_year_val})
        plot_data_list.append({'Category': str(ano_atual_referencia), 'Value': ytd_current_year_val})

        for _, row_month in df_metric_filtered[
            (df_metric_filtered['Period'].astype(str).str.isdigit()) &
            (pd.to_numeric(df_metric_filtered['Period'], errors='coerce').isin(range(1, 13))) &
            (df_metric_filtered['Year'] == ano_atual_referencia)
        ].iterrows():
            month_str = str(int(row_month['Period']))
            plot_data_list.append({'Category': month_str, 'Value': row_month['Value']})

        df_to_plot = pd.DataFrame(plot_data_list)
        df_to_plot = df_to_plot.groupby('Category')['Value'].sum().reset_index()

        meses_ordem_numeros = [str(i) for i in range(1, 13)]
        categories_order = [str(ano_anterior_referencia), str(ano_atual_referencia)] + meses_ordem_numeros

        df_to_plot['Category'] = pd.Categorical(df_to_plot['Category'], categories=categories_order, ordered=True)
        df_to_plot = df_to_plot.sort_values(by='Category').reset_index(drop=True)

        try:
            fig, ax = plt.subplots(figsize=(7, 4))
            # Paleta de cores: anos em roxo-escuro, meses em verde-teal
            n = len(df_to_plot)
            bar_colors = []
            for cat in df_to_plot['Category']:
                try:
                    int(cat)
                    if int(cat) > 100:   # é um ano (ex: 2025, 2026)
                        bar_colors.append('#4A4580')
                    else:                # é um mês (1-12)
                        bar_colors.append('#2D5F4C')
                except Exception:
                    bar_colors.append('#4A4580')

            x_pos = np.arange(len(df_to_plot))
            values = df_to_plot['Value'].fillna(0).values
            ax.bar(x_pos, values, color=bar_colors, width=0.65, zorder=2)
            ax.set_xticks(x_pos)
            ax.set_xticklabels(df_to_plot['Category'].tolist(), fontsize=8)
            ax.yaxis.grid(True, linestyle='--', alpha=0.4, zorder=0)
            ax.set_axisbelow(True)
            for spine in ['top','right']:
                ax.spines[spine].set_visible(False)
            ax.set_xlabel("")
            ax.set_ylabel("")

            final_rect_x_axes = 0.0
            final_rect_width_axes = 1.0
            rect_y_axes = 1.02
            rect_height_axes = 0.1

            rect = Rectangle((final_rect_x_axes, rect_y_axes), final_rect_width_axes, rect_height_axes,
                              transform=ax.transAxes, facecolor='#8c8c8c', edgecolor='#8c8c8c',
                              linewidth=1, clip_on=False, zorder=0)
            ax.add_patch(rect)

            center_y_in_rect_axes = rect_y_axes + (rect_height_axes / 2)
            ax.text(final_rect_x_axes + 0.01, center_y_in_rect_axes, metric_display_name,
                    transform=ax.transAxes, fontsize=12, weight='bold', color='white',
                    va='center', ha='left', zorder=1)

            if tem_emoji and tem_meta and pd.notna(valor_meta_para_grafico):
                # Emoji dinâmico: escolhe PNG baseado no status real vs meta
                ytd_val = df_metric_filtered[
                    df_metric_filtered['Period'] == str(ano_atual_referencia)
                ]['Value'].sum()
                if ytd_val == 0 or ytd_val <= valor_meta_para_grafico:
                    emoji_file = "emoji_green.png"
                elif ytd_val <= valor_meta_para_grafico * 1.05:
                    emoji_file = "emoji_yellow.png"
                else:
                    emoji_file = "emoji_red.png"

                if emoji_dir:
                    emoji_path = os.path.join(emoji_dir, emoji_file)
                    if os.path.exists(emoji_path):
                        from matplotlib.image import imread as mpl_imread
                        try:
                            img_arr = mpl_imread(emoji_path)
                            # Margem interna: 10% do retângulo no topo e base
                            margin   = rect_height_axes * 0.12
                            img_h    = rect_height_axes - 2 * margin
                            img_w    = img_h  # quadrado
                            inset_x  = 0.975 - img_w   # alinhado à direita com folga
                            inset_y  = rect_y_axes + margin
                            ax_inset = ax.inset_axes(
                                [inset_x, inset_y, img_w, img_h],
                                transform=ax.transAxes)
                            ax_inset.imshow(img_arr)
                            ax_inset.axis('off')
                            ax_inset.set_zorder(5)
                        except Exception as _e:
                            log_func(f"Aviso: emoji no gráfico: {_e}")

            max_value = df_to_plot['Value'].max() if not df_to_plot.empty else 0
            if pd.notna(max_value) and max_value > 0:
                y_upper_limit = max_value * 1.25
            else:
                y_upper_limit = 1.0
            ax.set_ylim(bottom=0, top=y_upper_limit)

            ANNOTATION_THRESHOLD = 0.5
            for xi, val in zip(x_pos, values):
                if val >= ANNOTATION_THRESHOLD:
                    if val == int(val):
                        fmt = f'{int(val):,}'.replace(',','X').replace('.', ',').replace('X','.')
                    else:
                        fmt = f'{val:,.1f}'.replace(',','X').replace('.', ',').replace('X','.')
                    ax.text(xi, val, fmt, ha='center', va='bottom',
                            fontsize=7, color='#1C1C1A')

            x_pos_annotation = len(df_to_plot) - 0.1

            if tem_meta and pd.notna(valor_meta_para_grafico):
                ax.axhline(y=valor_meta_para_grafico, color='red', linestyle='--', linewidth=1.5, label='Meta')
                formatted_meta_value = f'{valor_meta_para_grafico:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')

                if valor_meta_para_grafico == 0:
                    y_pos_annotation = ax.get_ylim()[1] * 0.05
                    if max_value == 0:
                        y_pos_annotation = ax.get_ylim()[1] * 0.1
                else:
                    y_pos_annotation = valor_meta_para_grafico + (ax.get_ylim()[1] * 0.02)
                    if y_pos_annotation > ax.get_ylim()[1] * 0.95:
                        y_pos_annotation = ax.get_ylim()[1] * 0.95

                ax.annotate(f'Meta: {formatted_meta_value}',
                            (x_pos_annotation, y_pos_annotation),
                            ha='right', va='bottom',
                            color='red', fontsize=9, weight='bold',
                            bbox=dict(boxstyle="round,pad=0.2", fc="yellow", ec="none", lw=0, alpha=0.8))

                log_func(f"Linha de meta {formatted_meta_value} adicionada para '{metric_original}' do cliente '{nome_cliente_para_grafico}'.")

            # Renomear categorias numéricas de mês para abreviação PT
            _MESES_ABREV = {
                '1':'Jan','2':'Fev','3':'Mar','4':'Abr','5':'Mai','6':'Jun',
                '7':'Jul','8':'Ago','9':'Set','10':'Out','11':'Nov','12':'Dez'
            }
            labels = [_MESES_ABREV.get(str(c), str(c))
                      for c in df_to_plot['Category'].tolist()]
            ax.set_xticklabels(labels, fontsize=8)
            plt.tight_layout(rect=[0, 0, 1, 0.95])
            nome_cliente_limpo_para_arquivo = clean_filename(nome_cliente_para_grafico)
            nome_arquivo_grafico = f"Grafico_{metric_original.replace(' ', '_').replace('/', '-')}_{nome_cliente_limpo_para_arquivo}.png"
            caminho_completo_grafico = os.path.join(pasta_graficos, nome_arquivo_grafico)
            plt.savefig(caminho_completo_grafico)
            plt.close('all')
            gc.collect()
            caminhos_graficos_gerados.append(caminho_completo_grafico)
        except Exception as e:
            log_func(f"Erro ao gerar gráfico '{metric_display_name}' para {nome_cliente_para_grafico}: {e}")
            traceback.print_exc()
            plt.close('all')

    return caminhos_graficos_gerados
