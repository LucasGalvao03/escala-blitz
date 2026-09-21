# Escala Blitz — DC-PI2

Site em Python (Flask) para visualizar e editar a escala de funcionários da
Blitz, terceirizada que presta serviço para a iMile Delivery na unidade
DC-PI2. Substitui a planilha original por um site simples, com os mesmos
dados (Julho, Agosto e Setembro de 2026 já importados).

## Como rodar localmente

```bash
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Depois abra http://localhost:5000 no navegador.

## Acesso (senha)

O site pede uma senha antes de deixar ver ou editar qualquer coisa. A senha
padrão é `blitz2026` — troque antes de usar de verdade, de dois jeitos:

- **Mais rápido:** edite a linha `LOGIN_PASSWORD` no topo do `app.py`.
- **Mais seguro:** defina a variável de ambiente `ESCALA_SENHA` antes de rodar
  o site, assim a senha não fica escrita no código:
  ```bash
  # Linux/Mac
  export ESCALA_SENHA="sua-senha-aqui"
  python app.py

  # Windows (cmd)
  set ESCALA_SENHA=sua-senha-aqui
  python app.py
  ```

Hoje é uma senha única compartilhada (não tem usuário por pessoa) — dá pra
compartilhar com o time. O botão **Sair**, no canto superior direito, encerra
a sessão.

## O que o site faz

- Uma aba por mês, e sub-abas por operação (quando o mês tem mais de uma,
  como "Operação DC" e "Operação DS").
- Tabela editável: clique em qualquer status do dia (T, DSR, Folga, Falta,
  Atestado, Abono ou vazio) para trocar, edite nome/turno/cargo direto na
  célula, adicione ou remova funcionários.
- Clique em **Salvar alterações** para gravar tudo de uma vez.
- Resumo automático por funcionário (folgas, faltas, atestados, abono, DSR,
  dias trabalhados, total de dias e status OK/NOT), recalculado sozinho.
- Botão **Exportar Excel (.xlsx)** gera uma planilha com a escala e o resumo
  do mês/operação que está na tela, já colorida por status.
- **+ Adicionar mês** cria um novo mês (com opção de copiar a lista de
  funcionários do último mês, deixando os status em branco).

## Onde ficam os dados

Tudo fica em `data/escala_data.json` — um arquivo de texto simples, sem
precisar de banco de dados. Para fazer backup, basta copiar esse arquivo.
Para editar os dados "na mão" (fora do site), é só abrir esse .json num
editor de texto — a estrutura é:

```json
{
  "months": [
    {
      "key": "2026-07",
      "label": "Julho 2026",
      "operations": [
        {
          "name": "Geral",
          "employees": [
            {"nome": "...", "turno": "T3", "cargo": "Auxiliar",
             "dias": {"2026-07-01": "DSR", "2026-07-02": "T", "...": "..."}}
          ]
        }
      ]
    }
  ]
}
```

## Hospedar de verdade (multiusuário, acesso pela internet)

Hoje o app roda localmente, para um usuário por vez, sem login. Para deixar
disponível para a equipe:

1. **Mais simples:** subir num serviço como Render, Railway ou PythonAnywhere
   (todos têm planos gratuitos/baratos para apps Flask pequenos).
2. Trocar `app.run(debug=True)` por um servidor de produção, ex.:
   `gunicorn app:app` (adicione `gunicorn` ao `requirements.txt`).
3. Se mais de uma pessoa for editar ao mesmo tempo, vale trocar o
   `escala_data.json` por um banco de dados de verdade (SQLite já resolve
   bem, sem precisar de servidor de banco separado).
4. Adicionar login (ex.: `flask-login`) se o site for exposto na internet,
   já que hoje qualquer pessoa com o link pode editar.

## Estrutura do projeto

```
escala_blitz_python/
├── app.py              # rotas e lógica (Flask)
├── requirements.txt
├── data/
│   └── escala_data.json    # os dados da escala
├── templates/
│   └── index.html          # template da página (Jinja2)
└── static/
    └── style.css            # visual do site
```
