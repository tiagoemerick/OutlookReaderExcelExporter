package com.viisi.br.util;

public final class Constants {

	private Constants() {
	}

	public static final class messages {
		public static final String FOLDER_NOT_FOUND = "A árvore de diretórios: 'Caixa de Entrada' não foi encontrada.";
	}

	public static final class outlook {
		public static final String INICIO_PASTAS_PARTICULARES = "Início de Pastas Particulares";
		public static final String ENTRADA = "Entrada";
		public static final String TITULO_EMAIL_PROCESSAMENTO = "[Teste - Processamento] - Teste";
		public static final String EMAIL_ORIGEM = "teste.teste@mail.com.br";
	}

	public static final class body {
		public static final String DATA_INCLUSAO_PROCESSAMENTO = "Data Inclusão:";
		public static final String DATA_FIM_PROCESSAMENTO = "Data Fim:";
		public static final String MENSAGEM_PROCESSAMENTO = "Mensagem:";
		public static final String JUNK = "Mensagem enviada automaticamente, não é necessário responder.";
	}

	public static final class lenght {
		public static final int HORA_DATA = 19;
	}

	public static final class excel {
		public static final String DIRETORIO_TEMPLATE = "/templateExcel";
		public static final String SEPARADOR_DIRETORIO_INTERNO_DO_PROJETO = "/";
		public static final String NOME_ARQUIVO_TEMPLATE = "TemplateExcel";
		public static final String EXTENSAO_ARQUIVO_TEMPLATE_EXCEL = ".xls";
		public static final String NOME_ABA = "Sheet1";
	}

}
