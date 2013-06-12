package com.viisi.br.example;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Vector;

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;

import com.viisi.br.outlook.PSTException;
import com.viisi.br.outlook.PSTFile;
import com.viisi.br.outlook.PSTFolder;
import com.viisi.br.outlook.PSTMessage;
import com.viisi.br.util.Constants;
import com.viisi.br.util.DataUtil;
import com.viisi.br.util.EscritorExcel;
import com.viisi.br.wrapper.DadosWrapper;

/**
 * Classe de teste para ler um diretório específico da caixa de entrada do
 * Outlook e gerar um arquivo excel. De acordo com título do email e remetente,
 * aplica determinada regra. Necessário ter o arquivo .psd com os emails.
 * 
 * @author tiago.emerick
 * @since 12/06/2013
 * 
 */
public class TestGui implements ActionListener {

	private PSTFile pstFile;
	private JFrame frame;

	public static void main(String[] args) throws PSTException, IOException {
		new TestGui();
	}

	@SuppressWarnings("deprecation")
	public TestGui() throws PSTException, IOException {
		frame = new JFrame("PSD Browser");

		String filename = buscarDiretorioPsd();
		try {
			String errorMessage = Constants.messages.FOLDER_NOT_FOUND;

			pstFile = new PSTFile(filename);
			if (pstFile.getRootFolder().hasSubfolders()) {
				Vector<PSTFolder> childFolders = pstFile.getRootFolder().getSubFolders();
				for (PSTFolder childFolder : childFolders) {
					if (childFolder.getDisplayName().equals(Constants.outlook.INICIO_PASTAS_PARTICULARES)) {
						Vector<PSTFolder> pastasParticulares = childFolder.getSubFolders();
						for (PSTFolder pasta : pastasParticulares) {
							if (pasta.getDisplayName().equalsIgnoreCase(Constants.outlook.ENTRADA)) {
								if (pasta.getContentCount() > 0) {
									int contentItemNumber = 1;
									errorMessage = null;

									List<DadosWrapper> listaDadosExportar = new ArrayList<DadosWrapper>();
									Date hoje = new Date();

									while (contentItemNumber <= pasta.getContentCount()) {
										PSTMessage email = (PSTMessage) pasta.getNextChild();
										if (email != null) {
											Calendar dataAtual = DataUtil.getDataComTimeZerado(hoje);
											Calendar dataRecebimento = DataUtil.getDataComTimeZerado(email.getMessageDeliveryTime());

											if (isRecebidoHoje(dataAtual, dataRecebimento)) {
												if (isEmailNotificacaoProcessamento(email)) {
													DadosWrapper dw = new DadosWrapper();

													final String mailBody = email.getBody();

													int begin = mailBody.indexOf(Constants.body.DATA_INCLUSAO_PROCESSAMENTO);
													int end = begin + Constants.body.DATA_INCLUSAO_PROCESSAMENTO.length();
													Date dataInclusao = new Date(mailBody.substring(end, end + Constants.lenght.HORA_DATA));
													dw.setDataInicio(DataUtil.formatarData(dataInclusao, DataUtil.FORMATO_DATA_HORA));

													begin = mailBody.indexOf(Constants.body.DATA_FIM_PROCESSAMENTO);
													end = begin + Constants.body.DATA_FIM_PROCESSAMENTO.length();
													Date dataFim = new Date(mailBody.substring(end, end + Constants.lenght.HORA_DATA));
													dw.setDataFim(DataUtil.formatarData(dataFim, DataUtil.FORMATO_DATA_HORA));

													String diferencaTempo = DataUtil.diferencaEmHoraMinutoSegundo(dataInclusao, dataFim);
													dw.setTempoProcessamento(diferencaTempo);

													int beginJunk = mailBody.indexOf(Constants.body.JUNK);
													begin = mailBody.indexOf(Constants.body.MENSAGEM_PROCESSAMENTO) + Constants.body.MENSAGEM_PROCESSAMENTO.length();
													end = beginJunk;
													String bodyTableContent = mailBody.substring(begin, end);

													String contentReplaced = bodyTableContent.replaceAll("\r", " ").trim();
													String[] tableContent = contentReplaced.split("\n");
													int numeroColuna = 1;
													int segundoEspaco = 1;
													int contador = 1;
													for (String column : tableContent) {
														if (numeroColuna == 1 && !column.trim().equals("")) {
															dw.setCooperativa(column.trim());
															numeroColuna++;
														} else if (numeroColuna == 2 && !column.trim().equals("")) {
															dw.setPeriodoContabil(column.trim());
															numeroColuna++;
														} else if (numeroColuna == 3 && !column.trim().equals("")) {
															dw.setSituacao(column.trim());
															numeroColuna++;

															if (contador == tableContent.length) {
																listaDadosExportar.add(dw);

																dw = new DadosWrapper();
																dw.setDataInicio(DataUtil.formatarData(dataInclusao, DataUtil.FORMATO_DATA_HORA));
																dw.setDataFim(DataUtil.formatarData(dataFim, DataUtil.FORMATO_DATA_HORA));
																dw.setTempoProcessamento(diferencaTempo);

																numeroColuna = 1;
															}
														} else if (numeroColuna == 4) {
															if (segundoEspaco == 2) {
																dw.setMensagemProcessamento(column.trim());
																listaDadosExportar.add(dw);

																dw = new DadosWrapper();
																dw.setDataInicio(DataUtil.formatarData(dataInclusao, DataUtil.FORMATO_DATA_HORA));
																dw.setDataFim(DataUtil.formatarData(dataFim, DataUtil.FORMATO_DATA_HORA));
																dw.setTempoProcessamento(diferencaTempo);

																numeroColuna = 1;
															} else {
																segundoEspaco++;
															}
														}
														contador++;
													}
												}
											}
										}
										contentItemNumber++;
									}
									if (!listaDadosExportar.isEmpty()) {
										StringBuilder sb = new StringBuilder("Teste - ");

										sb.append(DataUtil.getDia(hoje));
										sb.append(DataUtil.getMes(hoje));
										sb.append(DataUtil.getAno(hoje));
										sb.append(Constants.excel.EXTENSAO_ARQUIVO_TEMPLATE_EXCEL);

										EscritorExcel escritorExcel = new EscritorExcel(getCaminhoArquivoTemplateExcel(), sb.toString());
										escritorExcel.escreverInformacoes(Constants.excel.NOME_ABA, listaDadosExportar, getMetodosExportarExcel(), 1, 0);
										JOptionPane.showMessageDialog(null, "Gerado com sucesso!");
									} else {
										JOptionPane.showMessageDialog(null, "Não há nada a ser gerado hoje!");
									}
								} else {
									errorMessage = "A pasta '" + Constants.outlook.ENTRADA + "' está vazia.";
								}
							}
						}
					}
				}
			}
			if (errorMessage != null) {
				JOptionPane.showMessageDialog(null, errorMessage);
			}
		} catch (Exception err) {
			JOptionPane.showMessageDialog(null, "Erro na planilha: " + err.getMessage());
			err.printStackTrace();
		} finally {
			System.exit(1);
		}
	}

	private String buscarDiretorioPsd() throws IOException {
		String filename = null;

		File arquivo = new File("caminho_psd.txt");
		FileReader fr;
		try {
			fr = new FileReader(arquivo);
			BufferedReader br = new BufferedReader(fr);

			String string = br.readLine();
			while (string != null) {
				filename = string;
				break;
			}
			br.close();
			fr.close();
		} catch (FileNotFoundException e) {
			filename = requestToChooseFile();

			FileWriter fw = new FileWriter(arquivo);
			BufferedWriter xbw = new BufferedWriter(fw);
			xbw.write(filename);
			xbw.flush();
			xbw.close();
		} catch (IOException e) {
			JOptionPane.showMessageDialog(null, "Erro na planilha: " + e.getMessage());
			e.printStackTrace();
		}
		return filename;
	}

	public static List<String> getMetodosExportarExcel() {
		List<String> resultado = new ArrayList<String>();
		resultado.addAll(new ArrayList<String>(Arrays.asList("getDataInicio", "getDataFim", "getTempoProcessamento", "getCooperativa", "getPeriodoContabil", "getSituacao", "getMensagemProcessamento")));

		return resultado;
	}

	private String getCaminhoArquivoTemplateExcel() {
		String diretorioTemplate = Constants.excel.DIRETORIO_TEMPLATE;
		String separadorDiretorio = Constants.excel.SEPARADOR_DIRETORIO_INTERNO_DO_PROJETO;
		String extensaoArquivoTemplate = Constants.excel.EXTENSAO_ARQUIVO_TEMPLATE_EXCEL;

		return diretorioTemplate + separadorDiretorio + Constants.excel.NOME_ARQUIVO_TEMPLATE + extensaoArquivoTemplate;
	}

	private boolean isEmailNotificacaoProcessamento(PSTMessage email) {
		return email != null && email.getSenderEmailAddress().equalsIgnoreCase(Constants.outlook.EMAIL_ORIGEM) && email.getSubject().equalsIgnoreCase(Constants.outlook.TITULO_EMAIL_PROCESSAMENTO);
	}

	private boolean isRecebidoHoje(Calendar dataAtual, Calendar dataRecebimento) {
		return dataRecebimento.compareTo(dataAtual) == 0;
	}

	private String requestToChooseFile() throws IOException {
		JFileChooser chooser = new JFileChooser();
		if (chooser.showOpenDialog(frame) == JFileChooser.APPROVE_OPTION) {
		} else {
			System.exit(0);
		}

		String filename = chooser.getSelectedFile().getCanonicalPath();

		return filename;
	}

	@Override
	public void actionPerformed(ActionEvent e) {
	}

}
