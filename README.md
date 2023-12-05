# WhatsApp_Bot_Excel
Imagine conseguir melhorar o desempenho da sua empresa na manutenção e fidelização dos clientes de um estúdio fotográfico, além de propectar novos clientes com incríveis feedbacks recebidos? Tudo isso possível com mensagens pelo WhatsApp na hora certa e da forma certa!
Essa demanda foi realizada por um Estúdio Fotografico, que solicitou um BOT WhatsApp capaz de automatizar o envio de mensagens para clientes na data de aniversário, para feedback 15 dias após entrega do produto, para oferecer descontos aos clientes fidelizados.

Obtive os dados dos clientes na planilha,  de acordo com a demanda solicitada, usando o Openpyxl.
	- Nome da cliente
	- Número de telefone (continha DDD na planilha mas não tinha o código do país, necessário para enviar mensagens através do link do whatsApp Web)
	- Data de aniversário (padrão utilizado pela cliente: %d/%m)
	- Data da última compra (padrão utilizado pela cliente: %d/%m/%Y)
	- Quantidade de compras efetuadas (número inteiro)

Iniciei o driver Selenium para o navegador Edge, utilizado pela cliente (Webdriver.Edge( ))

Antes de executar qualquer comando, solicitei abertura do navegador no link do WhatsApp Web, e então utilizei o loop While para aguardar a cliente logar no WhatsApp Web, para então continuar o código.

Preparei várias condições para atender a demanda da cliente. Para cada condição, foi criado uma mensagem personalizada. Como todas as condições iriam realizar as mesmas ações no WhatsApp Web, utilizando o Selenium, criei uma função para otimizar o código. As condições foram:
	- SE a data atual for a data do aniversário daquela cliente - disparar uma linda mensagem de aniversário, oferecendo desconto.
	- SE a data atual for igual a 15 dias após a data da última compra - disparar a mensagem solicitando Feedback à cliente.
	- SE a data atual for igual a 6 meses após a data da última compra - disparar a mensagem de remarketing.
	- SE a quantidade de compras for igual ou superior a 3, ou menor que 6 e o ultimo contato feito para fidelização for superior a 3 meses - disparar a mensagem para cliente fiel básico
	- SE a quantidade de compras for igual ou superior a 6 e o ultimo contato feito para fidelização for superior a 3 meses - disparar a mensagem para cliente fiel
Iterei então sobre a planilha Excel para obter os valores de cada um dos dados indicados mais acima. 
Durante a iteração, as condições acima então foram executadas para cada conjunto de dados obtido. 

Sendo satisfeita a condição, a mensagem então é criada e a função enviar_mensagem é então disparada, com a mensagem como argumento.

A função enviar_mensagem(mensagem) executa as ações:
	- Criação do link padrão para WhatsApp Web, no formato: https://web.whatsapp.com/send?phone=55{telefone}&text={quote(mensagem)}, com o uso do quote, da biblioteca urllib.parse para padronizar o texto.
	- Abrir o link, no navegador já aberto, e logado.
	- Encontrar o XPATH (find_element, By.XPATH) que localiza o botão de enviar mensagem, e então clicar nele, usando o comando click( )
	- Executar uma pausa de 15 segundos para evitar possiveis bloqueios envio de mensagens em massa (SPAM)

## Dificuldades e Soluções:

- Datas utilizadas em padrões diferentes: As datas existentes na planilha da cliente seguiam o padrão %d/%m/%Y. Precisei adequar a data do dia (datetime.today( )) para ficar compatível com as datas utilizadas pela cliente na planilha. Feito através do strftime('%d/%m/%Y').

- Na resolução da demanda da cliente, para enviar mensagem de fidelização às clientes que compraram mais de 3 vezes, o BOT inicialmente criado enviava mensagem toda vez que o aplicativo era aberto, ocasionando em um SPAM. Para tratar esse problema, criei uma nova coluna na planilha da cliente, chamada ultimo contato feito para fidelização. Cada vez que o BOT envia mensagem à cliente, registra a data atual no campo da planilha. Assim, pude incluir na condição IF, para envio de mensagem por fidelização, a condição de apenas enviar mensagem SE o último contato para fidelização for superior a 3 meses.

- O Aplicativo apresentava erro por não encontrar o XPATH indicado. Percebi que o erro acontecia enquanto o navegador ainda atualizava a página para incluir a mensagem para o cliente. Resolvi o problema adicionando um tempo de espera de 15 segundos (sleep da biblioteca Time), entre a abertura do navegador com o link e o código para localizar e clicar no elemento que contém o XPATH indicado.

- Se a cliente sem querer fechar o navegador aberto, que já estava logado, e então o aplicativo, durante a execução da função enviar_mensagem, abrir novamente o navegador, agora com o link contendo o numero do telefone e a mensagem para envio, o navegador solicitará novamente que associe o dispositivo. Isso gera então o mesmo erro do XPATH acima, pois mesmo tratando o problema anterior, 15 segundos podem não ser suficientes para a cliente logar novamente e o navegador atualizar com a mensagem pronta. Para resolver isso, coloquei aqui o mesmo loop (while) que criei no inicio do aplicativo, para aguardar enquanto a cliente não loga. Coloquei essa linha de código logo após o sleep de 15 segundos, e antes do código para encontrar o XPATH e clicar no botão para enviar a mensagem. 

- Caso o número do cliente não exista, ou caso o WhatsApp demore na renderização da tela, ou ainda caso o WhatsApp esteja fora do ar... enfim, caso haja algum problema e não seja possível o BOT localizar o caminho XPATH com o botão de envio da mensagem, para evitar a quebra do código, inclui todas as condições dentro de um Try / Except. O Except então armazena, num arquivo erros.csv, o nome e o telefone dos clientes que não foram possíveis executar a ação.

- Como a cliente pediu para executar esse código na inicialização do seu computador, de forma automática, tive que incluir um sleep  inicial para aguardar o computador ligar e a cliente logar. Coloquei um sleep de 180 segundos mas que precisa adequar, dependendo do computador e de como é feito o login nele.

- ## Tecnologias Usadas:

- - Selenium (para acessar e automatizar o WhatsApp Web)
- Openpyxl (acesso aos dados na planilha da cliente) - Poderia usar o Pandas, mas como já tinha feito outros projetos usando Pandas, preferi usar o Openpyxl para conhecer a biblioteca.
- Datetime (para padronizar as datas de acordo com o formato de data usado pela cliente na planilha)
- Time (Usei essa biblioteca para forçar o aplicativo a parar e aguardar alguns segundos, evitando possiveis bloqueios por envio de mensagens em massa.)
