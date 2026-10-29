# Confirmação de presença em festas

Página simples para os convidados confirmarem presença, com a imagem do convite no topo.
O convidado informa o nome e quantos **acompanhantes adultos** e **crianças (até 12 anos)**
vêm junto. No painel de admin você cria quantos eventos quiser, cada um com seu link, e
gera a **lista em ordem alfabética para a portaria** (imprimir ou copiar para o WhatsApp).

Roda de graça no **Google Apps Script**, preso a uma Planilha Google sua. Os convidados
não precisam de login, e as respostas aparecem também na planilha.

## Instalação (uma vez só, ~10 minutos)

1. Crie uma planilha nova em <https://sheets.new> e dê um nome a ela (ex.: *Confirmações de festas*).
2. Na planilha, abra **Extensões → Apps Script**.
3. No editor que abrir:
   - Apague o conteúdo de `Código.gs` e cole o conteúdo de [`apps-script/Codigo.gs`](apps-script/Codigo.gs).
   - **Troque a senha** na linha `const SENHA_ADMIN = 'troque-esta-senha';`.
   - Clique em **+ → HTML** e crie três arquivos com estes nomes exatos (sem `.html`), colando o conteúdo de cada um:
     - `Convidado` ← [`apps-script/Convidado.html`](apps-script/Convidado.html)
     - `Admin` ← [`apps-script/Admin.html`](apps-script/Admin.html)
     - `Estilo` ← [`apps-script/Estilo.html`](apps-script/Estilo.html)
   - Salve (ícone de disquete).
4. Clique em **Implantar → Nova implantação**:
   - Tipo: **App da Web**
   - Executar como: **Eu**
   - Quem pode acessar: **Qualquer pessoa**
   - Clique em **Implantar** e autorize o acesso (o Google vai avisar que o app "não foi verificado":
     clique em *Avançado → Acessar (não seguro)*. É normal, porque o app é seu).
5. Copie a **URL do app da Web** (termina em `/exec`).

## Uso

- **Painel de admin:** abra a URL com `?admin` no fim, ex.: `https://script.google.com/macros/s/XXXX/exec?admin`.
  Guarde nos favoritos.
- **Novo evento:** *+ Novo evento* → nome, mensagem (opcional), imagem do convite → *Salvar*.
  Aparece o **link para os convidados** (botões *Copiar* e *Enviar no WhatsApp*).
- **Lista da portaria:** abra o evento → *Gerar lista da portaria*. Sai em ordem alfabética, com
  adultos (já contando o convidado), crianças, total e uma caixinha "Chegou" para o porteiro marcar.
  Botões *Imprimir* e *Copiar texto*.
- **Encerrar confirmações:** desmarque *Aceitando confirmações* e salve.
- Se a mesma pessoa confirmar de novo com o mesmo nome, a resposta dela é **atualizada**, não duplicada.
- Para apagar uma resposta, use o **×** na tabela de confirmações.

## Onde ficam os dados

- Aba **Eventos** e aba **Confirmações** da planilha (criadas sozinhas no primeiro uso).
- Imagens dos convites: pasta *Convites de festa (confirmação)* no seu Google Drive. Ela continua
  privada, porque a página mostra a imagem pelo próprio script.

## Atualizar o código depois

Cole a nova versão dos arquivos no editor e vá em **Implantar → Gerenciar implantações → ✏️ →
Versão: Nova versão → Implantar**. Assim o link continua o mesmo.

## Observações

- O Google mostra uma faixa cinza no topo ("Este aplicativo foi criado por um usuário do Google
  Apps Script"). Não dá para tirar no plano gratuito.
- A primeira abertura do link pode levar 2–3 segundos.
