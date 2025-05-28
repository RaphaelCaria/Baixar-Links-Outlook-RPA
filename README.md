🤖 RPA - Download Automático de Anexos via Outlook e OneDrive
Automação desenvolvida em Python que interage com o Outlook para localizar e baixar anexos de e-mails específicos cujos arquivos estão armazenados no OneDrive. Ideal para cenários corporativos onde há necessidade recorrente de obter documentos automaticamente por e-mail.

🔧 Tecnologias Utilizadas
Python 3.x

pywin32 (integração com Outlook via COM)

os, shutil (manipulação de arquivos)

datetime (filtros por data)

Integração com OneDrive (via caminho local sincronizado)

📁 Estrutura do Projeto
main.py: script principal de execução da automação.

config.py: parâmetros de configuração (e-mail, pasta destino, filtros).

logs/: armazena registros de execução e erros (se aplicável).

downloads/: pasta onde os arquivos baixados são salvos automaticamente.

▶️ Como Executar
Certifique-se de que o Outlook está instalado e com conta logada.

Sincronize previamente o OneDrive com a máquina local.

Ajuste os parâmetros em config.py conforme necessário.

Execute o script com: python main.py

Os arquivos serão salvos automaticamente na pasta PDFS_BAIXADOS.

🚀 Funcionalidades
Leitura automática de e-mails no Outlook.

Identificação de e-mails com anexos hospedados no OneDrive.

Download dos arquivos em pasta local.

Possibilidade de aplicar filtros por data, assunto ou remetente.

Pode ser agendado via Task Scheduler do Windows para execuções recorrentes.
