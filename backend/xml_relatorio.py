import xmltodict
import pandas as pd
from pathlib import Path


class Relatorio:
    """Gera relatórios em XLSX a partir de arquivos XML de andamentos."""

    arquivos_xml = [
        r'C:\Users\antho\OneDrive\Desktop\projeto xml_data-frame\backend\pasta_XML\TEST2.xml' #altere aqui o path do xml para teste.
    ]

    responsaveis = {
        '49840': 'ROSANE',
        '30756': 'LUCIA',
        '34309': 'JOSIANE',
        '32917': 'LUARA',
        '26964': 'NATALIA',
        '26936': 'IVILIANE',
        '29770': 'ANTHONY'
    }

    def responsavel_evento(self, codigo):
        """Converte o código de solicitante no nome da responsável."""
        if not codigo:
            return "Sem responsavel"
        for chave, nome in self.responsaveis.items():
            if codigo.startswith(chave):
                return codigo.replace(chave, nome)
        return codigo

    def __init__(self):
        self.lista_clientes = []

    def processar_xml(self):
        """Lê os XMLs configurados e popula `lista_clientes` com os dados."""
        for arquivo in self.arquivos_xml:
            with open(arquivo, encoding="utf8") as f:
                conteudo_xml = f.read()

            conteudo_xml_dicionario = xmltodict.parse(conteudo_xml)
            andamentos = conteudo_xml_dicionario['andamentos']['andamento']

            if isinstance(andamentos, dict):
                andamentos = [andamentos]

            for andamento in andamentos:
                processo = andamento['processo']

                cliente = {
                    'PJ': processo.get('PJ', ''),
                    'Evento': processo.get('perfil', ''),
                    'Conformidade': '',
                    'Responsável': self.responsavel_evento(andamento.get('solicitado_por')),
                    'data_evento': andamento['data_evento']
                }

                self.lista_clientes.append(cliente)            
    
    def relatorio_xlsx(self):
        """Gera ou atualiza o XLSX com as linhas coletadas do XML."""
        df_novos = pd.DataFrame(self.lista_clientes)

        df_novos['data_evento'] = pd.to_datetime(df_novos['data_evento'], format='%d/%m/%Y', errors='coerce').dt.date
        #altere aqui path da planilha xlsx para teste.
        output_path = Path(r'C:\Users\antho\OneDrive\Desktop\projeto xml_data-frame\backend\pasta_XLSX\andamentos.xlsx')

        if output_path.exists():
            df_existente = pd.read_excel(output_path)

            if 'data_evento' in df_existente.columns:
                df_existente['data_evento'] = pd.to_datetime(df_existente['data_evento'], errors='coerce').dt.date
            
        
            df_atualizado = pd.concat([df_existente, df_novos], ignore_index=True)
            df_atualizado = df_atualizado.sort_values(by='data_evento')
            df_atualizado = df_atualizado.drop_duplicates(
                subset=['PJ', 'Evento', 'data_evento'],
                keep='first'
            )
        else:
            df_atualizado = df_novos

       
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_atualizado.to_excel(writer, index=False, sheet_name='Sheet1')
        
            worksheet = writer.sheets['Sheet1']
            col_data_idx = df_atualizado.columns.get_loc('data_evento') + 1

            for row in range(2, len(df_atualizado) + 2):
                cell = worksheet.cell(row=row, column=col_data_idx)
                if cell.value:
                    cell.number_format = 'DD/MM/YYYY'

        print(f"Relatório atualizado com sucesso em {output_path}.")


if __name__ == "__main__":
    relatorio = Relatorio()
    relatorio.processar_xml()
    relatorio.relatorio_xlsx()
