import pandas as pd
import os

def gerar_site_estoque():
    # Define o caminho para a pasta onde o script está salvo
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    # Ajustado para o nome do arquivo excel que você mencionou
    nome_arquivo = 'ESTOQUE200126.xlsx' 
    caminho_completo = os.path.join(diretorio_atual, nome_arquivo)
    
    if not os.path.exists(caminho_completo):
        print(f"ERRO: Arquivo {nome_arquivo} não encontrado em {diretorio_atual}")
        return

    try:
        # Carrega o arquivo Excel (Certifique-se de ter instalado: pip install openpyxl)
        df = pd.read_excel(caminho_completo)
        
        # Remove espaços extras nos nomes das colunas para evitar erros de busca
        df.columns = [str(col).strip() for col in df.columns]
        
        print("Colunas encontradas no seu Excel:", df.columns.tolist())
        
    except Exception as e:
        print(f"Erro ao ler o Excel: {e}")
        return

    # Início da estrutura HTML
    html_base = f"""
    <!DOCTYPE html>
    <html lang="pt-br">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>G-LAB PEPTIDES - Estoque Oficial</title>
        <style>
            body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #f0f4f8; margin: 0; padding: 20px; }}
            .container {{ max-width: 1100px; margin: auto; background: white; padding: 30px; border-radius: 15px; box-shadow: 0 10px 25px rgba(0,0,0,0.1); }}
            .header {{ text-align: center; margin-bottom: 30px; padding-bottom: 20px; border-bottom: 2px solid #e1e8ed; }}
            .logo {{ max-width: 280px; height: auto; }}
            .search-container {{ margin-bottom: 25px; }}
            .search-box {{ width: 100%; padding: 15px; border: 2px solid #004a99; border-radius: 10px; font-size: 16px; outline: none; transition: 0.3s; }}
            table {{ width: 100%; border-collapse: collapse; }}
            th {{ background: #004a99; color: white; padding: 18px 12px; text-align: left; text-transform: uppercase; font-size: 13px; }}
            td {{ padding: 15px 12px; border-bottom: 1px solid #f0f0f0; font-size: 14px; color: #333; }}
            tr:hover {{ background: #f8faff; }}
            .status-disponivel {{ color: #27ae60; font-weight: bold; background: #eafaf1; padding: 6px 15px; border-radius: 20px; font-size: 12px; border: 1px solid #27ae60; }}
            .status-espera {{ color: #d35400; font-weight: bold; background: #fef5e7; padding: 6px 15px; border-radius: 20px; font-size: 12px; border: 1px solid #d35400; }}
            .sku {{ font-family: monospace; color: #7f8c8d; background: #f4f6f7; padding: 3px 6px; border-radius: 4px; }}
            .preco {{ font-weight: 700; color: #004a99; font-size: 15px; }}
            .footer {{ text-align: center; margin-top: 40px; color: #95a5a6; font-size: 13px; border-top: 1px solid #eee; padding-top: 20px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <img src="1.png" alt="G-LAB PEPTIDES" class="logo">
                <h2 style="color: #004a99; margin-top: 10px;">Inventário de Peptídeos</h2>
            </div>
            
            <div class="search-container">
                <input type="text" id="inputBusca" class="search-box" onkeyup="filtrar()" placeholder="Pesquisar por produto ou código SKU...">
            </div>

            <table id="tabelaEstoque">
                <thead>
                    <tr>
                        <th>Produto</th>
                        <th>Vol/Medida</th>
                        <th>SKU (Código)</th>
                        <th>Status</th>
                        <th>Preço Unitário</th>
                    </tr>
                </thead>
                <tbody>
    """

    # Preenchimento dinâmico baseado nos dados fornecidos 
    for _, row in df.iterrows():
        # Captura os dados usando os nomes exatos das colunas do seu arquivo 
        produto = str(row.get('PRODUTO', '---'))
        vol = str(row.get('VOL', '')).replace('nan', '')
        medida = str(row.get('MEDIDA', '')).replace('nan', '')
        sku = str(row.get('SKU', '---'))
        
        # Lógica de Status: Verifica "Estoque" (comum em Excel) ou "ESTOQUE" 
        status_val = str(row.get('Estoque', row.get('ESTOQUE', 'EM ESPERA'))).strip().upper()
        
        # Captura o preço exato 
        preco = str(row.get('Preço (R$)', row.get('PREÇO (R$)', 'Consulte')))

        classe_status = "status-disponivel" if "DISPON" in status_val else "status-espera"
        
        html_base += f"""
                    <tr>
                        <td><strong>{produto}</strong></td>
                        <td>{vol} {medida}</td>
                        <td><span class="sku">{sku}</span></td>
                        <td><span class="{classe_status}">{status_val}</span></td>
                        <td class="preco">{preco}</td>
                    </tr>
        """

    html_base += f"""
                </tbody>
            </table>
            <div class="footer">
                <p><strong>G-LAB PEPTIDES</strong> [cite: 3]</p>
                <p>Website: <a href="https://glabpeptides.com/" style="color:#004a99;">glabpeptides.com</a> [cite: 4, 8] | Contato: +1 (774) 622-2523 [cite: 5, 9]</p>
            </div>
        </div>

        <script>
        function filtrar() {{
            var input, filter, table, tr, td, i, txtValue;
            input = document.getElementById("inputBusca");
            filter = input.value.toUpperCase();
            table = document.getElementById("tabelaEstoque");
            tr = table.getElementsByTagName("tr");

            for (i = 1; i < tr.length; i++) {{
                var mostrar = false;
                var colunas = tr[i].getElementsByTagName("td");
                for (var j = 0; j < colunas.length; j++) {{
                    if (colunas[j]) {{
                        txtValue = colunas[j].textContent || colunas[j].innerText;
                        if (txtValue.toUpperCase().indexOf(filter) > -1) {{
                            mostrar = true;
                            break;
                        }}
                    }}
                }}
                tr[i].style.display = mostrar ? "" : "none";
            }}
        }}
        </script>
    </body>
    </html>
    """

    caminho_html = os.path.join(diretorio_atual, "index.html")
    with open(caminho_html, "w", encoding="utf-8") as f:
        f.write(html_base)
    
    print(f"SUCESSO: O arquivo 'index.html' foi gerado com {len(df)} itens processados!")

if __name__ == "__main__":
    gerar_site_estoque()