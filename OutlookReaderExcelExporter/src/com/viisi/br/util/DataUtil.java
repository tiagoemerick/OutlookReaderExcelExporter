package com.viisi.br.util;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.Locale;

import org.apache.commons.lang3.StringUtils;

public class DataUtil {

	public static final String FORMATO_MES_ANO = "MM/yyyy";
	public static final String FORMATO_DIA_MES_ANO = "dd/MM/yyyy";
	public static final String FORMATO_DATA_HORA_COM_TRACO = "dd-MM-yyyy-HHmm";
	public static final String FORMATO_ANO_MES_DIA = "yyyy-MM-dd";
	public static final String FORMATO_ANO_MES_DIA_HORA_MIN_SEG = "yyyy-MM-dd HH:mm:ss";
	public static final String FORMATO_DATA_HORA = "dd/MM/yyyy hh:mm";
	public static final String FORMATO_ANO_MES_DIA_SEM_HIFEN = "yyyyMMdd";
	public static final String FORMATO_DATA_HORA_SEGUNDOS = "MM/dd/yyyy HH:mm:ss";
	public static final String FORMATO_DBA_HORA_SEGUNDOS = "dd/MM/yyyy HH:mm:ss";
	public static final String FORMATO_HORA_MINUTO_SEGUNDO_SEM_HIFEN = "HHmmss";
	public static final String FORMATO_HORA_MINUTO_SEGUNDO = "HH:mm:ss";

	public static final String FORMATO_DIA_MES_ANO_SEM_HIFEN = "ddMMyyyy";

	private static final int HORAS_NO_DIA = 24;
	private static final int MINUTOS_NA_HORA = 60;
	private static final int SEGUNDOS_NO_MINUTO = 60;
	private static final int MILISSEGUNDOS_NO_SEGUNDO = 1000;

	private static final Locale LOCALE = new Locale("pt", "BR");

	/**
	 * Recupera um formatador para datas, usando o formato especificado.
	 * 
	 * @param formato
	 *            O formata a ser utilizado pelo formatador.
	 * @return Um <code>SimpleDateFormat</code> que será o formatador para
	 *         datas.
	 */
	public static SimpleDateFormat getFormatadorData(String formato) {
		return new SimpleDateFormat(formato, LOCALE);
	}

	/**
	 * Formata a data especificada de acordo com o formado passado.
	 * 
	 * @param data
	 *            A data a ser formatada.
	 * @param formato
	 *            O formato a ser seguido.
	 * @return Uma <code>String</code> representando a data formatada.
	 */
	public static String formatarData(Date data, String formato) {
		return data != null ? getFormatadorData(formato).format(data) : null;
	}

	/**
	 * Formata uma data para o formato MM/yyyy. O método recebe uma string como
	 * parâmetro representando a data no formata yyyyMM.
	 * 
	 * @param data
	 *            A string que representa a data a ser formatada.
	 * 
	 * @return Um string formatada no formato MM/yyyy.
	 */
	public static String formataDataParaMesAno(String data) {
		if (data.length() >= 6) {
			String mes = data.substring(4, 6);
			String ano = data.substring(0, 4);

			return formataDataParaMesAno(mes, ano);
		}

		return "";
	}

	/**
	 * Retorna a data com hora/minuto/segundo/millisegundo/hora do dia zerado
	 * 15/07/2013 00:00:00
	 * 
	 * @return Calendar
	 */
	public static Calendar getDataComTimeZerado(Date dataPassada) {
		Calendar dataAtual = GregorianCalendar.getInstance();
		dataAtual.setTime(dataPassada);
		dataAtual.set(Calendar.HOUR, 0);
		dataAtual.set(Calendar.MINUTE, 0);
		dataAtual.set(Calendar.SECOND, 0);
		dataAtual.set(Calendar.MILLISECOND, 0);
		dataAtual.set(Calendar.HOUR_OF_DAY, 0);
		// dataAtual.set(Calendar.DAY_OF_MONTH, 11);
		return dataAtual;
	}

	/**
	 * Formata uma data para o formato yyyy/MM. O método recebe uma string como
	 * parâmetro representando a data no formata MMyyyy.
	 * 
	 * @param data
	 *            A string que representa a data a ser formatada.
	 * 
	 * @return Um string formatada no formato MM/yyyy.
	 */
	public static String formataDataParaAnoMes(String data) {
		if (data.length() >= 6) {
			String ano = data.substring(0, 4);
			String mes = data.substring(4, 6);

			return formataDataParaMesAno(ano, mes);
		}

		return "";
	}

	/**
	 * Formata uma data para o formato MM/yyyy. O método recebe duas strings
	 * como parâmetro, uma representando o mês e a outra representando o ano.
	 * 
	 * Se a string que representa o mês tiver tamanho igual a 1, ou seja, o mês
	 * está entre janeiro (1) e setembro (9), é adicionado um zero à esquerda da
	 * string para que o seu valor seja 01, 02, 03, etc.
	 * 
	 * @param mes
	 *            A string que representa o mês.
	 * @param ano
	 *            A string que representa o ano.
	 * 
	 * @return Um string formatada no formato MM/yyyy.
	 */
	public static String formataDataParaMesAno(String mesParametro, String ano) {
		String mes = mesParametro;
		if (mes != null && ano != null) {
			mes = mes.length() == 1 ? "0" + mes : mes;
			return mes + "/" + ano;
		}

		return "";
	}

	/**
	 * Subtrai a quantidade de meses especificado de uma determinada data.
	 * 
	 * @param data
	 *            A data que será subtraida.
	 * @param meses
	 *            A quantidade de meses a subtrair.
	 * @return Um objeto <code>Date</code> representando a data especificada
	 *         subtraida de 6 meses.
	 */
	public static Date subtrairMeses(Date data, int meses) {

		Calendar calendario = Calendar.getInstance();
		calendario.setTime(data);

		calendario.add(Calendar.MONTH, -meses);

		return calendario.getTime();
	}

	/**
	 * Devolve o dia da data informada em inteiro
	 * 
	 * @author RenatoFStefanini
	 * @param data
	 * @return dia
	 */
	public static int getDia(Date data) {
		String dia = formatarData(data, "dd");
		return Integer.parseInt(dia);
	}

	/**
	 * Devolve o mês da data informada em inteiro
	 * 
	 * @author RenatoFStefanini
	 * @param data
	 * @return dia
	 */
	public static int getMes(Date data) {
		String mes = formatarData(data, "MM");
		return Integer.parseInt(mes);
	}

	/**
	 * Devolve o ano da data informada em inteiro
	 * 
	 * @author RenatoFStefanini
	 * @param data
	 * @return dia
	 */
	public static int getAno(Date data) {
		String ano = formatarData(data, "yyyy");
		return Integer.parseInt(ano);
	}

	public static Date getData(String dataString, String formatoDaData) throws ParseException {
		SimpleDateFormat dateFormat = new SimpleDateFormat(formatoDaData);
		dateFormat.setLenient(false);
		return dateFormat.parse(dataString);
	}

	/**
	 * Calcula o tempo em horas, minutos e segundos entre duas datas.
	 * 
	 * @param data1
	 *            A data inicial.
	 * @param data2
	 *            A data final.
	 * @return uma string, no formato hh:mm:ss, com o tempo decorrido
	 */
	public static long calcularDiasEntreDatas(Date data1, Date data2) {
		long diferenca = 0;

		if (data2.after(data1)) {
			diferenca = data2.getTime() - data1.getTime();
		} else {
			diferenca = data1.getTime() - data2.getTime();
		}

		long horas = ((diferenca / MILISSEGUNDOS_NO_SEGUNDO) / SEGUNDOS_NO_MINUTO) / MINUTOS_NA_HORA;

		return (horas / HORAS_NO_DIA);
	}

	/**
	 * Retorna a diferença em 00:00:00
	 * 
	 * @param inicio
	 * @param fim
	 * @return String
	 */
	public static String diferencaEmHoraMinutoSegundo(Date inicio, Date fim) {
		long difMilli = fim.getTime() - inicio.getTime();
		int timeInSeconds = (int) difMilli / 1000;
		int hours, minutes, seconds;
		hours = timeInSeconds / 3600;
		timeInSeconds = timeInSeconds - (hours * 3600);
		minutes = timeInSeconds / 60;
		timeInSeconds = timeInSeconds - (minutes * 60);
		seconds = timeInSeconds;

		return StringUtils.leftPad(String.valueOf(hours), 2, "0") + ":" + StringUtils.leftPad(String.valueOf(minutes), 2, "0") + ":" + StringUtils.leftPad(String.valueOf(seconds), 2, "0");
	}

	/**
	 * Adiciona a quantidade de dias especificado a uma determinada data.
	 * 
	 * @param data
	 *            A data que será adicionada.
	 * @param meses
	 *            A quantidade de dias a adicionar.
	 * @return Um objeto <code>Date</code> representando a data especificada
	 *         adicionada da quantidade de dias.
	 */
	public static Date adicionarDias(Date data, int dias) {
		Calendar calendario = Calendar.getInstance();
		calendario.setTime(data);

		calendario.add(Calendar.DATE, dias);

		return calendario.getTime();
	}

	/**
	 * Devolve o mês da data informada em String
	 * 
	 * @author RenatoFStefanini
	 * @param data
	 * @return dia
	 */
	public static String getMesString(Date data) {
		String mes = formatarData(data, "MM");
		if (mes.length() == 1) {
			mes += "0";
		}
		return mes;
	}

	/**
	 * Devolve o ano da data informada em String
	 * 
	 * @author RenatoFStefanini
	 * @param data
	 * @return dia
	 */
	public static String getAnoString(Date data) {
		String ano = formatarData(data, "yyyy");

		return ano;
	}

	/**
	 * Devolve o ultimo dia do mes especificado
	 * 
	 * @author clebiosctis
	 * @param data
	 * @return data
	 */
	public static Date retornaUltimoDia(Date data) {
		Calendar calendario = Calendar.getInstance();
		calendario.setTime(data);
		calendario.set(Calendar.DATE, calendario.getMaximum(Calendar.DATE) - 1);

		return calendario.getTime();
	}

	/**
	 * Formata uma data para o formato yyyyMMdd. O método recebe uma string como
	 * parâmetro representando a data no formata dd/MM/yyyy.
	 * 
	 * @param dataNascimento
	 *            A string que representa a data a ser formatada.
	 * 
	 * @return Uma string formatada no formato yyyyMMdd.
	 */
	public static String formataDataParaAnoMesDia(String dataNascimento) {
		if (dataNascimento.length() >= 10) {
			String ano = dataNascimento.substring(6);
			String mes = dataNascimento.substring(3, 5);
			String dia = dataNascimento.substring(0, 2);

			return ano + mes + dia;
		}
		return "";
	}

	public static Date removerHora(Date data) {
		Calendar calendar = Calendar.getInstance();
		calendar.setLenient(false);
		calendar.setTime(data);
		calendar.set(Calendar.HOUR_OF_DAY, 0);
		calendar.set(Calendar.MINUTE, 0);
		calendar.set(Calendar.SECOND, 0);
		calendar.set(Calendar.MILLISECOND, 0);

		return calendar.getTime();
	}
}