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
 * Classe responsável por escrever num arquivo excel os dados passados como
 * parâmetros pelos métodos.
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
	 *            exportação.
	 * @param caminhoArquivoDestino
	 *            O caminho para onde deverá ser mandado o arquivo criado com os
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
	 * Verifica a existência de um arquivo com o mesmo nome do que o que
	 * armazenará a lista dos dados.
	 * 
	 * Caso já exista um arquivo com o mesmo nome, o arquivo antigo é deletado
	 * para que não haja duplicidade de informações.
	 */
	private void limparArquivosExistentes() {
		File arquivoDestino = new File(caminhoArquivoDestino);
		if (arquivoDestino.exists()) {
			arquivoDestino.delete();
		}
	}

	/**
	 * Escreve os dados do relatório de liquidação financeira na aba
	 * especificada horizontalmente.
	 * 
	 * @throws VisaNetException
	 *             Lançada caso ocorra algum erro na escrita do arquivo.
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

			// Obtém a planilha do arquivo excel onde os dados serão escritos
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
			throw new Exception("O arquivo " + caminhoArquivoDestino + " não foi encontrado para escrita.");
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
	 *             Lançada caso não exista uma planilha no arquivo com o nome
	 *             especificado.
	 */
	private HSSFSheet getPlanilha(HSSFWorkbook arquivoExcel, String nomePlanilha) throws Exception {
		HSSFSheet planilha = arquivoExcel.getSheet(nomePlanilha);

		if (planilha == null) {
			throw new Exception("No arquivo " + caminhoArquivoTemplate + " não existe uma planilha com o nome " + nomePlanilha + ".");
		}
		return planilha;
	}

	/**
	 * Recupera a informação retornada pela execução do método especificado do
	 * objeto passado como parâmetro.
	 * 
	 * @param objeto
	 *            O objeto que terá o método invocado.
	 * @param nomeMetodo
	 *            O nome do método a ser invocado.
	 * 
	 * @return O retorno do método em formato de <code>String</code>. Caso o a
	 *         invocação do método resulte em <code>null</code>, é retornado uma
	 *         <code>String</code> vazia.
	 * @throws Exception
	 *             Lançada caso ocorra algum erro na invocação do método.
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
			throw new Exception("Caminho do arquivo template ou destino do relatório XLS é nulo.");
		}

		try {
			criarCopiaDoArquivoTemplate();
		} catch (IOException e) {
			throw new Exception("Não foi possível criar a cópia do arquivo template do relatório XLS. Erro: " + e.getMessage());
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
			System.out.println("NÃO FUNCIONA " + this.getClass().getClassLoader());
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
