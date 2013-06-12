package com.viisi.br.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

/**
 * Classe respons�vel por escrever num arquivo excel os dados passados como
 * par�metros pelos m�todos.
 * 
 * @author Stefanini IT Solutions
 */
public class EscritorExcel {

	private String caminhoArquivoTemplate;
	private String caminhoArquivoDestino;
	private String caminhoArquivoTemp;

	@SuppressWarnings("unused")
	private int linhaInicial;
	@SuppressWarnings("unused")
	private int colunaInicial;

	@SuppressWarnings("unused")
	private HSSFWorkbook workbook;

	/**
	 * Construtor da classe.
	 * 
	 * @param caminhoArquivoTemplate
	 *            O caminho de onde se encontra o arquivo de template para a
	 *            exporta��o.
	 * @param caminhoArquivoDestino
	 *            O caminho para onde dever� ser mandado o arquivo criado com os
	 *            dados.
	 * @throws Exception
	 */
	public EscritorExcel(String caminhoArquivoTemplate, String caminhoArquivoDestino) throws Exception {
		this.caminhoArquivoTemplate = caminhoArquivoTemplate;
		this.caminhoArquivoDestino = caminhoArquivoDestino;

		limparArquivosExistentes();

		criarCopiaDoTemplate();
	}

	/**
	 * Verifica a exist�ncia de um arquivo com o mesmo nome do que o que
	 * armazenar� a lista dos dados.
	 * 
	 * Caso j� exista um arquivo com o mesmo nome, o arquivo antigo � deletado
	 * para que n�o haja duplicidade de informa��es.
	 */
	private void limparArquivosExistentes() {
		File arquivoDestino = new File(caminhoArquivoDestino);
		if (arquivoDestino.exists()) {
			arquivoDestino.delete();
		}
	}

	/**
	 * Escreve os dados do relat�rio de liquida��o financeira na aba
	 * especificada horizontalmente.
	 * 
	 * @throws VisaNetException
	 *             Lan�ada caso ocorra algum erro na escrita do arquivo.
	 */
	public void escreverInformacoes(String nomeAba, List<? extends Object> dados, List<String> metodos, int linhaInicial, int colunaInicial) throws Exception {
		HSSFCell celula = null;

		this.linhaInicial = linhaInicial;
		this.colunaInicial = colunaInicial;

		try {
			String inExecuteDir = System.getProperty("user.dir");
			String fullFileName = inExecuteDir + "\\" + caminhoArquivoDestino.substring(0, caminhoArquivoDestino.indexOf(Constants.excel.EXTENSAO_ARQUIVO_TEMPLATE_EXCEL)) + Constants.excel.EXTENSAO_ARQUIVO_TEMPLATE_EXCEL;
			File file = new File(fullFileName);

			// Abre arquivo
			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(caminhoArquivoTemp));

			// cria os estilos das linhas
			HSSFCellStyle estiloAlinhamentoEsquerdaBranco = workbook.createCellStyle();
			estiloAlinhamentoEsquerdaBranco.setAlignment(HSSFCellStyle.ALIGN_LEFT);

			HSSFCellStyle estiloAlinhamentoEsquerdaCinza = workbook.createCellStyle();
			estiloAlinhamentoEsquerdaCinza.setAlignment(HSSFCellStyle.ALIGN_LEFT);
			estiloAlinhamentoEsquerdaCinza.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			estiloAlinhamentoEsquerdaCinza.setFillForegroundColor(new HSSFColor.GREY_25_PERCENT().getIndex());

			HSSFCellStyle estiloAlinhamentoCentroBranco = workbook.createCellStyle();
			estiloAlinhamentoCentroBranco.setAlignment(HSSFCellStyle.ALIGN_CENTER);

			HSSFCellStyle estiloAlinhamentoCentroCinza = workbook.createCellStyle();
			estiloAlinhamentoCentroCinza.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			estiloAlinhamentoCentroCinza.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			estiloAlinhamentoCentroCinza.setFillForegroundColor(new HSSFColor.GREY_25_PERCENT().getIndex());

			// Obt�m a planilha do arquivo excel onde os dados ser�o escritos
			HSSFSheet planilha = getPlanilha(workbook, nomeAba);

			int numLinha = linhaInicial;

			for (Object objeto : dados) {
				if (numLinha <= 65535) {
					int numColuna = colunaInicial;
					for (String metodo : metodos) {
						HSSFRichTextString informacao = new HSSFRichTextString(getInformacao(objeto, metodo));
						inserirInformacoesCelula(celula, planilha, informacao, estiloAlinhamentoCentroCinza, estiloAlinhamentoCentroBranco, numLinha, numColuna, false);

						numColuna++;
					}
				}
				numLinha++;
			}
			FileOutputStream fileOut = new FileOutputStream(file);

			workbook.write(fileOut);
			fileOut.close();
		} catch (FileNotFoundException e) {
			throw new Exception("O arquivo " + caminhoArquivoDestino + " n�o foi encontrado para escrita.");
		} catch (Exception e) {
			limparArquivosExistentes();
			throw new Exception("Problemas na escrita do arquivo: " + caminhoArquivoDestino + ". ERRO: " + e.getMessage());
		}
	}

	/**
	 * Recupera a planilha especificada do arquivo excel.
	 * 
	 * @param arquivoExcel
	 *            O arquivo excel.
	 * @param nomePlanilha
	 *            O nome da planilha a ser recuperada.
	 * 
	 * @return A planilha requerida.
	 * @throws Exception
	 *             Lan�ada caso n�o exista uma planilha no arquivo com o nome
	 *             especificado.
	 */
	private HSSFSheet getPlanilha(HSSFWorkbook arquivoExcel, String nomePlanilha) throws Exception {
		HSSFSheet planilha = arquivoExcel.getSheet(nomePlanilha);

		if (planilha == null) {
			throw new Exception("No arquivo " + caminhoArquivoTemplate + " n�o existe uma planilha com o nome " + nomePlanilha + ".");
		}
		return planilha;
	}

	/**
	 * Recupera a informa��o retornada pela execu��o do m�todo especificado do
	 * objeto passado como par�metro.
	 * 
	 * @param objeto
	 *            O objeto que ter� o m�todo invocado.
	 * @param nomeMetodo
	 *            O nome do m�todo a ser invocado.
	 * 
	 * @return O retorno do m�todo em formato de <code>String</code>. Caso o a
	 *         invoca��o do m�todo resulte em <code>null</code>, � retornado uma
	 *         <code>String</code> vazia.
	 * @throws Exception
	 *             Lan�ada caso ocorra algum erro na invoca��o do m�todo.
	 */
	private String getInformacao(Object objeto, String nomeMetodo) throws Exception {
		String valor = null;
		Method metodo = objeto.getClass().getMethod(nomeMetodo, new Class[0]);
		Object retorno = metodo.invoke(objeto, new Object[0]);
		valor = retorno != null ? retorno.toString() : new String();

		return valor;
	}

	private void criarCopiaDoTemplate() throws Exception {
		if (caminhoArquivoTemplate == null || caminhoArquivoDestino == null) {
			throw new Exception("Caminho do arquivo template ou destino do relat�rio XLS � nulo.");
		}

		try {
			criarCopiaDoArquivoTemplate();
		} catch (IOException e) {
			throw new Exception("N�o foi poss�vel criar a c�pia do arquivo template do relat�rio XLS. Erro: " + e.getMessage());
		}
	}

	private void criarCopiaDoArquivoTemplate() throws IOException {
		InputStream inputStream = null;
		try {
			String caminhoArquivo = caminhoArquivoTemplate.substring(1, caminhoArquivoTemplate.length());
			inputStream = this.getClass().getClassLoader().getResourceAsStream(caminhoArquivo);
			if (inputStream == null) {
				inputStream = this.getClass().getResourceAsStream(caminhoArquivoTemplate);
			}
		} catch (Exception e) {
			inputStream = this.getClass().getResourceAsStream(caminhoArquivoTemplate);
			System.out.println("N�O FUNCIONA " + this.getClass().getClassLoader());
			e.printStackTrace();

		}
		caminhoArquivoTemp = caminhoArquivoDestino;

		File file = new File(caminhoArquivoTemp.substring(0, caminhoArquivoTemp.indexOf(Constants.excel.EXTENSAO_ARQUIVO_TEMPLATE_EXCEL)) + Constants.excel.EXTENSAO_ARQUIVO_TEMPLATE_EXCEL);

		OutputStream outputStream = new FileOutputStream(file);

		int len;
		byte[] buf = new byte[1024];
		while ((len = inputStream.read(buf)) > 0) {
			outputStream.write(buf, 0, len);
		}

		inputStream.close();
		outputStream.close();
		caminhoArquivoTemp = file.getAbsolutePath();

		workbook = new HSSFWorkbook(new FileInputStream(caminhoArquivoTemp));
	}

	@SuppressWarnings("deprecation")
	private void inserirInformacoesCelula(HSSFCell celula, HSSFSheet planilha, HSSFRichTextString informacao, HSSFCellStyle estiloAlinhamentoCentroCinza, HSSFCellStyle estiloAlinhamentoCentroBranco, int numLinha, int numColuna, boolean isAdicional) {
		HSSFRow linha;
		// obtem linha
		linha = planilha.getRow(numLinha);

		if (linha == null) {
			linha = planilha.createRow(numLinha);
			for (int i = 0; i < 100; i++) {
				linha.createCell((short) i);
			}
		}
		celula = linha.getCell((short) numColuna);
		if (celula == null) {
			linha.createCell((short) numColuna);
			celula = linha.getCell((short) numColuna);
		}
		if (celula != null) {
			celula.setCellValue(informacao);
		} else {
			linha = planilha.createRow(numLinha);
			linha.createCell((short) numColuna);
			celula = linha.getCell((short) numColuna);
			celula.setCellValue(informacao);
		}
	}

}
