import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.cluster import KMeans, DBSCAN, AgglomerativeClustering
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA

# ------------------------
# Funções para selecionar arquivo e planilha
# ------------------------
def pick_excel_file():
    """
    Abre uma caixa de diálogo para o usuário selecionar um arquivo Excel.
    Retorna um objeto ExcelFile (xls) ou None se o usuário cancelar.
    """
    try:
        import tkinter as tk
        from tkinter import filedialog
        import os
        
        root = tk.Tk()
        root.withdraw()
        
        file_path = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm *.xlsb"), ("All files", "*.*")]
        )
        if not file_path:
            print("Nenhum arquivo foi selecionado. Encerrando.")
            return None
        
        xls = pd.ExcelFile(file_path)
        print(f"\nArquivo selecionado: {os.path.basename(file_path)}")
        return xls
    
    except Exception as e:
        print(f"Erro ao abrir arquivo: {e}")
        return None

def ask_for_sheet_name(xls):
    sheets = xls.sheet_names
    while True:
        print("\nPlanilhas disponíveis:", sheets)
        sheet_name = input("Digite o nome da planilha a ser utilizada (ou 'q' para sair): ").strip()
        if sheet_name.lower() == 'q':
            return None
        if sheet_name in sheets:
            return sheet_name
        else:
            print(f"A planilha '{sheet_name}' não foi encontrada. Tente novamente ou digite 'q' para sair.")

def read_excel_interactively(xls, sheet_name):
    try:
        df_preview = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    except Exception as e:
        print(f"Erro ao carregar a planilha '{sheet_name}': {e}")
        return None
    
    print("\nPré-visualização das 10 primeiras linhas (SEM cabeçalho):")
    print(df_preview.head(10))
    
    while True:
        header_row = input("\nDigite o número (0-based) da linha que contém o cabeçalho "
                           "(ou 'n' se não houver, ou 'q' para sair): ").strip().lower()
        
        if header_row == 'q':
            return None
        
        if header_row == 'n':
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                break
            except Exception as e:
                print(f"Erro ao carregar planilha sem cabeçalho: {e}")
                return None
        else:
            try:
                row_num = int(header_row)
                df = pd.read_excel(xls, sheet_name=sheet_name, header=row_num)
                break
            except ValueError:
                print("Valor inválido. Tente novamente ou 'q' para sair.")
            except Exception as e:
                print(f"Erro ao carregar planilha com cabeçalho na linha {header_row}: {e}")
                return None
    
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
    
    print("\nExemplo das 5 primeiras linhas do DataFrame resultante:")
    print(df.head(5))
    print(f"Shape final: {df.shape}")
    return df

# ------------------------
# Função para filtragem por data
# ------------------------
def filter_by_date(df):
    """
    Procura por colunas que possam conter datas.
    Se houver candidatas, pergunta ao usuário se deseja filtrar por data.
    Caso sim, pergunta qual coluna usar e as datas inicial e final.
    Retorna o DataFrame filtrado.
    """
    candidates = []
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            candidates.append(col)
        else:
            if df[col].dtype == object:
                sample = df[col].dropna().head(10)
                if sample.empty:
                    continue
                converted = pd.to_datetime(sample, errors='coerce')
                if converted.notna().sum() / len(sample) > 0.8:
                    candidates.append(col)
    candidates = list(set(candidates))
    
    if not candidates:
        print("Não foram encontradas colunas de datas.")
        return df
    else:
        print("\nColunas candidatas a datas encontradas:")
        for i, col in enumerate(candidates, 1):
            print(f"{i}. {col}")
        while True:
            filter_choice = input("Deseja filtrar o dataset por um intervalo de datas? (s/n): ").strip().lower()
            if filter_choice in ['s', 'n']:
                break
            else:
                print("Digite 's' para sim ou 'n' para não.")
        if filter_choice == 'n':
            return df
        else:
            # Se houver mais de uma coluna candidata, o usuário pode escolher pelo nome ou número
            date_col = None
            if len(candidates) > 1:
                while True:
                    chosen = input("Digite o nome ou número (conforme listado) da coluna que deseja usar para filtrar: ").strip()
                    if chosen.isdigit():
                        idx = int(chosen) - 1
                        if 0 <= idx < len(candidates):
                            date_col = candidates[idx]
                            break
                        else:
                            print("Número inválido. Tente novamente.")
                    elif chosen in candidates:
                        date_col = chosen
                        break
                    else:
                        print("Entrada inválida. Escolha entre as opções acima.")
            else:
                date_col = candidates[0]
                print(f"Usando a coluna {date_col} para filtragem de data.")
            try:
                df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
            except Exception as e:
                print(f"Erro ao converter coluna {date_col} para data: {e}")
                return df
            
            while True:
                start_date_str = input("Digite a data inicial (YYYY-MM-DD): ").strip()
                try:
                    start_date = pd.to_datetime(start_date_str, format="%Y-%m-%d")
                    break
                except Exception:
                    print("Formato inválido. Tente novamente.")
            while True:
                end_date_str = input("Digite a data final (YYYY-MM-DD): ").strip()
                try:
                    end_date = pd.to_datetime(end_date_str, format="%Y-%m-%d")
                    break
                except Exception:
                    print("Formato inválido. Tente novamente.")
            
            original_rows = df.shape[0]
            df = df[(df[date_col] >= start_date) & (df[date_col] <= end_date)]
            filtered_rows = df.shape[0]
            print(f"Filtragem aplicada: {filtered_rows} linhas restantes de {original_rows}.")
            return df

# ------------------------
# Função para selecionar colunas (aceitando números ou nomes)
# ------------------------
def ask_for_columns(df):
    """
    Exibe as colunas disponíveis e permite ao usuário escolher-as por número e/ou nome.
    Retorna uma lista de nomes de colunas.
    """
    col_names = list(df.columns)
    while True:
        print("\nColunas disponíveis:")
        for i, col in enumerate(col_names, 1):
            print(f"{i}. {col}")
        
        cols_input = input("Digite os nomes ou números das colunas para clusterização, separados por vírgula (ou 'q' para sair): ").strip()
        if cols_input.lower() == 'q':
            return None
        
        tokens = [t.strip() for t in cols_input.split(",")]
        selected = []
        for token in tokens:
            if token.isdigit():
                idx = int(token) - 1
                if 0 <= idx < len(col_names):
                    selected.append(col_names[idx])
                else:
                    print(f"Número inválido: {token}")
            else:
                if token in col_names:
                    selected.append(token)
                else:
                    print(f"Coluna não encontrada: {token}")
        if selected:
            return list(dict.fromkeys(selected))  # remove duplicatas
        else:
            print("Nenhuma coluna válida selecionada. Tente novamente.")

# ------------------------
# Outras funções de conversão, seleção de algoritmos e execução dos mesmos
# ------------------------
def convert_columns_to_numeric(df, columns):
    for col in columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    data = df[columns].dropna()
    return data

def ask_for_algorithms():
    """
    Permite que o usuário escolha os algoritmos por número ou nome.
    Opções:
      1 - kmeans
      2 - dbscan
      3 - agglomerative
    """
    algo_options = {
        "1": "kmeans",
        "2": "dbscan",
        "3": "agglomerative",
        "kmeans": "kmeans",
        "dbscan": "dbscan",
        "agglomerative": "agglomerative"
    }
    while True:
        print("\nOpções de algoritmos:")
        print("1. kmeans")
        print("2. dbscan")
        print("3. agglomerative")
        algo_input = input("Digite os algoritmos desejados (por número ou nome, separados por vírgula) ou 'q' para sair: ").strip().lower()
        if algo_input == 'q':
            return None
        tokens = [t.strip() for t in algo_input.split(",")]
        chosen = []
        for token in tokens:
            if token in algo_options:
                chosen.append(algo_options[token])
            else:
                print(f"Opção inválida: {token}")
        if chosen:
            return list(dict.fromkeys(chosen))
        else:
            print("Nenhuma opção válida selecionada. Tente novamente.")

def run_kmeans(data_scaled, df, data_index):
    while True:
        n_clusters = input("Para KMeans, informe o número de clusters (ou 'c' para cancelar): ").strip()
        if n_clusters.lower() == 'c':
            print("Cancelado KMeans.")
            return None
        try:
            n_clusters = int(n_clusters)
            kmeans = KMeans(n_clusters=n_clusters, random_state=42)
            labels_kmeans = kmeans.fit_predict(data_scaled)
            df.loc[data_index, "cluster_kmeans"] = labels_kmeans
            print("KMeans clustering concluído.")
            return labels_kmeans
        except ValueError:
            print("Valor inválido. Tente novamente.")

def run_dbscan(data_scaled, df, data_index):
    while True:
        eps = input("Para DBSCAN, informe o valor de eps (ou 'c' para cancelar): ").strip()
        if eps.lower() == 'c':
            print("Cancelado DBSCAN.")
            return None
        try:
            eps = float(eps)
            min_samples = int(input("Para DBSCAN, informe o valor de min_samples: ").strip())
            dbscan = DBSCAN(eps=eps, min_samples=min_samples)
            labels_dbscan = dbscan.fit_predict(data_scaled)
            df.loc[data_index, "cluster_dbscan"] = labels_dbscan
            print("DBSCAN clustering concluído.")
            return labels_dbscan
        except ValueError:
            print("Valor inválido. Tente novamente.")

def run_agglomerative(data_scaled, df, data_index):
    while True:
        n_clusters_ag = input("Para Agglomerative, informe o número de clusters (ou 'c' para cancelar): ").strip()
        if n_clusters_ag.lower() == 'c':
            print("Cancelado Agglomerative.")
            return None
        try:
            n_clusters_ag = int(n_clusters_ag)
        except ValueError:
            print("Número inválido. Tente novamente.")
            continue
        
        # Opções de linkage: permitir números ou nomes
        linkage_options = {
            "1": "ward",
            "2": "complete",
            "3": "average",
            "4": "single",
            "ward": "ward",
            "complete": "complete",
            "average": "average",
            "single": "single"
        }
        print("\nOpções de linkage para Agglomerative:")
        print("1. ward")
        print("2. complete")
        print("3. average")
        print("4. single")
        linkage_in = input("Informe o tipo de linkage (por número ou nome): ").strip().lower()
        if linkage_in not in linkage_options:
            print("Linkage inválido, tente novamente.")
            continue
        linkage = linkage_options[linkage_in]
        
        try:
            agglo = AgglomerativeClustering(n_clusters=n_clusters_ag, linkage=linkage)
            labels_agg = agglo.fit_predict(data_scaled)
            df.loc[data_index, "cluster_agglomerative"] = labels_agg
            print("Agglomerative clustering concluído.")
            return labels_agg
        except ValueError:
            print("Valor inválido. Tente novamente.")
        except MemoryError as e:
            print(f"Erro de memória: {e}. Tente reduzir o tamanho do dataset ou usar outro algoritmo.")
            return None

# ------------------------
# Função de plotagem que utiliza os nomes das colunas
# ------------------------
def plot_clusters(data_scaled, labels_dict, selected_columns):
    n_features = data_scaled.shape[1]
    
    if n_features == 1:
        x_values = np.arange(len(data_scaled))
        data_1d = data_scaled[:, 0]
        for algo_name, labels in labels_dict.items():
            if labels is None:
                continue
            plt.figure(figsize=(8, 4))
            plt.title(f"Clusters (1D) - {algo_name} (Feature: {selected_columns[0]})")
            plt.scatter(x_values, data_1d, c=labels, cmap="viridis", s=50)
            plt.xlabel("Índice do dado")
            plt.ylabel(selected_columns[0])
            plt.colorbar(label="Cluster")
            plt.show()
    elif n_features == 2 and len(selected_columns) == 2:
        for algo_name, labels in labels_dict.items():
            if labels is None:
                continue
            plt.figure(figsize=(8, 6))
            plt.title(f"Clusters - {algo_name} ({selected_columns[0]} vs {selected_columns[1]})")
            scatter = plt.scatter(data_scaled[:, 0], data_scaled[:, 1], c=labels, cmap="viridis", s=50)
            plt.xlabel(selected_columns[0])
            plt.ylabel(selected_columns[1])
            plt.colorbar(scatter, label="Cluster")
            plt.show()
    else:
        if n_features > 2:
            pca = PCA(n_components=2)
            data_2d = pca.fit_transform(data_scaled)
        else:
            data_2d = data_scaled
        for algo_name, labels in labels_dict.items():
            if labels is None:
                continue
            plt.figure(figsize=(8, 6))
            plt.title(f"Clusters - {algo_name} (PCA of features: {', '.join(selected_columns)})")
            scatter = plt.scatter(data_2d[:, 0], data_2d[:, 1], c=labels, cmap="viridis", s=50)
            plt.xlabel("Componente 1")
            plt.ylabel("Componente 2")
            plt.colorbar(scatter, label="Cluster")
            plt.show()

# ------------------------
# Função para salvar arquivo usando caixa de diálogo
# ------------------------
def save_file_dialog(df):
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        output_path = filedialog.asksaveasfilename(
            title="Salvar arquivo Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if output_path:
            df.to_excel(output_path, index=False)
            print(f"Arquivo salvo como {output_path}")
        else:
            print("Salvamento cancelado.")
    except Exception as e:
        print(f"Erro ao salvar o arquivo: {e}")

# ------------------------
# Função principal
# ------------------------
def main():
    print("=== Leitura de Excel via caixa de diálogo e Clusterização ===")
    
    xls = pick_excel_file()
    if xls is None:
        print("Saindo do programa.")
        return
    
    sheet_name = ask_for_sheet_name(xls)
    if sheet_name is None:
        print("Saindo do programa.")
        return
    
    df = read_excel_interactively(xls, sheet_name)
    if df is None:
        print("Saindo do programa.")
        return

    df = filter_by_date(df)
    
    columns = ask_for_columns(df)
    if not columns:
        print("Saindo do programa.")
        return
    
    data = convert_columns_to_numeric(df, columns)
    if data.empty:
        print("Não há dados válidos nas colunas selecionadas. Encerrando.")
        return
    
    scaler = StandardScaler()
    data_scaled = scaler.fit_transform(data)
    
    algorithms = ask_for_algorithms()
    if not algorithms:
        print("Nenhum algoritmo selecionado. Encerrando.")
        return
    
    labels_dict = {}
    for algo in algorithms:
        if algo == "kmeans":
            labels_dict["KMeans"] = run_kmeans(data_scaled, df, data.index)
        elif algo == "dbscan":
            labels_dict["DBSCAN"] = run_dbscan(data_scaled, df, data.index)
        elif algo == "agglomerative":
            labels_dict["Agglomerative"] = run_agglomerative(data_scaled, df, data.index)
    
    plot_clusters(data_scaled, labels_dict, columns)
    
    while True:
        salvar = input("Deseja salvar o DataFrame com os clusters em um novo arquivo Excel? (s/n): ").strip().lower()
        if salvar in ['s', 'n']:
            break
        else:
            print("Opção inválida. Digite 's' ou 'n'.")
    if salvar == 's':
        save_file_dialog(df)
    else:
        print("Arquivo não será salvo.")
    
    print("\n=== Fim do programa ===")

if __name__ == "__main__":
    main()
