import java.io.BufferedInputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.sql.Date;

import javax.faces.bean.ManagedBean;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.util.PDFTextStripper;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookType;

import com.sun.prism.paint.Color;
import com.sun.xml.internal.ws.util.StringUtils;
 


@ManagedBean
public class EFPROCESSUAL {
	
	private static GetSetCEF getSetCEF = new GetSetCEF();
	private static EficienciaFinanceiraPROCESSUAL EficienciaFinanceiraBB = new EficienciaFinanceiraPROCESSUAL();
	private static String banco = "PROCESSUAL";  // PROCESSUAL
	

  	static String caminho = "";
  	static String excelBB = "";
	
	
  	static ArrayList<String> arrayPublicoCJExisteUnica = new ArrayList<String> ();
  	
  	
  	static File fo = new File(excelBB);

	//public static void main(String[] args) throws IOException {	
	public static void init() throws IOException, InterruptedException {
		fo = new File(excelBB);
		visualizarArquivos();
		
	}

	
	  public static void visualizarArquivos() throws IOException, InterruptedException {
		  


		  	String arquivoPDF = "";	
		  	String ret = "";
		  	String excluidos  = "";
		  
		  	File file = new File(caminho);
			File afile[] = file.listFiles();
			int i = 0;
			
			for (int j = afile.length; i < j; i++) {
				File arquivos = afile[i];
				arquivoPDF = caminho+arquivos.getName();
				//System.out.println(arquivos.getName());
				
				if(arquivos.getName().indexOf ("pdf") >= 0) {
				
					//Recebe e Le Texto do PDF
					String texto = extraiTextoDoPDF(arquivoPDF);
					
					//Recebe o Conteudo do PDF e coloca em Array com quebra de Linha
					String linhas[] = texto.split("\n");
					
			        					        			
					//Trata o PDF lendo linha a linha do array
					 ret = trataPDF(linhas);
					
					 
					 if(arquivoPDF.indexOf ("DS_Store") <= 0 ) //Se arquivo for diferente de arquivo de sistema que nao precisa ser analizado ou se o valor é zero mas nao igual ao mes de consulta
					 {
						 lerExcel("a");
						 //getSetCEF.setPorcentagem((i*100)/afile.length);
						 //System.out.println(" STATUS: ["+getSetCEF.getPorcentagem()+" %]");
					 }	
				}
			}
			
			
			getSetCEF.setFimArquivo("FINALIZADO COM SUCESSO");
			System.out.println("FINALIZADO COM SUCESSO");
		}
	
  

    		//EXTRAI OS TEXTOS DE DENTRO DO PDF
		  public static String extraiTextoDoPDF(String caminho) {
			  
			  if(caminho.indexOf ("DS_Store") >= 0) {
				  return "NAOePDF";
			  }
			  
			    PDDocument pdfDocument = null;
			    try {
			      pdfDocument = PDDocument.load(caminho);
			      PDFTextStripper stripper = new PDFTextStripper();
			      String texto = stripper.getText(pdfDocument);
			      return texto;
			    } catch (IOException e) {
			      throw new RuntimeException(e);
			    } finally {
			      if (pdfDocument != null) try {
			        pdfDocument.close();
			      } catch (IOException e) {
			        throw new RuntimeException(e);
			      }
			    }
		  }
			  
		  
		  //RECEBE O CONTEUDO DO EXCEL PARA TRATAMENTO DAS INFORMACOES
		  public static String trataPDF(String[] linhas) 
		  {			  
			  String ret = "";
			  boolean EXTRATOFALTACAMPOS = false;
			  
			  //System.out.println("PASSO 1");

				  if(banco.equals("PROCESSUAL")) // Banco do Brasil
				  {
					  for (int i = 0; i < linhas.length; i++) 
					  {
						  
						  ret = linhas[i];
						  //System.out.println(""+ret + " ["+i+"]");
						  
							
						  //PEGANDO A VARA E COMARCA
						  if(ret.contains("Detalhes do Processo") )
				          {
				           		ret = linhas[i];
				           		
				           		String[] Valor = null;
				           		Valor =	linhas[i].split("-");
								ret = Valor[3];
				            		
				           		Valor =	ret.split("ª");
								ret = Valor[0];

								//SETANDO A VARA
								getSetCEF.setVara(ret+"ª");  
								
				            	//	System.out.println(" - - - - - - - - - - - - - - >> VARA : ["+getSetCEF.getVara()+"]");

				          
								//Peggando a COMARCA
				            		ret = Valor[1];
				            		ret = ret.replace(" Vara do Trabalho de ", "");
				            		ret = ret.replace("Vara do Trabalho de", "");
				            		ret = ret.replace(" Vara do Trabalho do ", "");
				            		ret = ret.replace("Vara do Trabalho do", "");

				            		ret = ret.replace("(", "");
								ret = ret.replace(")", "");
								
				            		//SETANDO A COMARCA
				            		getSetCEF.setComarca(ret);
					       //     	System.out.println(" - - - - - - - - - - - - - - >> COMARCA : ["+getSetCEF.getComarca()+"]");

				          
				          }
						 
						  //PEGANDO O NUMERO DO PROCESSO
						  if(ret.contains("Processo PJe") )
				          {
				           		
							  	ret = linhas[i];
							  	
							  	String[] Valor = null;
							  	
								Valor =	linhas[i].split(":");
								
								ret = Valor[1];
								ret = tirarAlfabeto(ret);
								ret = ret.replace("(", "");
								ret = ret.replace(")", "");
								ret = ret.trim();
								
								
								if(ret.substring(0,1).equals("-"))
								{
									ret = ret.substring(1,ret.length());
								}
				            		
				            		//SETANDO A CONTA JUDICIAL
				            		getSetCEF.setNumeroProcesso(ret);// para gravar no excel
				            		getSetCEF.setNumeroProcessoAuxiliar(ret.replaceAll("[^0-9]", "")); // Auxiliar que pega apenas numero
				            		
				            		
//				            		System.out.println(" - - - - - - - - - - - - - - >> NUMERO PROCESSO: ["+getSetCEF.getNumeroProcesso()+"]");
				            		//System.out.println(" - - - - - - - - - - - - - - >> NUMERO PROCESSO: ["+getSetCEF.getNumeroProcessoAuxiliar()+"]");
				          }
						  
						  
						  //AUTOR
						  if(ret.contains("AUTOR(S)") || ret.contains("RECLAMANTE(S)"))
				          {
				           		
							  	ret = linhas[i];
							  	
							  	String[] Valor = null;
							  	
								Valor =	linhas[i].split(":");
								
								ret = Valor[1];
				            		
				            		//SETANDO O AUTOR
				            		getSetCEF.setAutor(ret);// para gravar no excel
				            		
				            		//System.out.println(" - - - - - - - - - - - - - - >>  AUTOR: ["+getSetCEF.getAutor()+"]");
				          }
						  
						  //REU
						  if(ret.contains("RÉU(S)") || ret.contains("RECLAMADO(S)"))
				          {
				           		
							  	ret = linhas[i];
							  	
							  	String[] Valor = null;
							  	
								Valor =	linhas[i].split(":");
								
								ret = Valor[1];
								
								ret = ret.replace("(+ 1)", "");
								ret = ret.replace("(+ 2)", "");
								ret = ret.replace("(+ 3)", "");
								ret = ret.replace("(+ 4)", "");
								ret = ret.replace("(+ 5)", "");
								ret = ret.replace("(+ 6)", "");
								ret = ret.replace("(+ 7)", "");
								ret = ret.replace("(+ 8)", "");
								ret = ret.replace("(+ 9)", "");
								ret = ret.replace("(+ 10)", "");
								
				            		
				            		//SETANDO REU
				            		getSetCEF.setReu(ret);// para gravar no excel
				            		
				            		//System.out.println(" - - - - - - - - - - - - - - >>  REU: ["+getSetCEF.getReu()+"]");
				          }   
					  }
				  }	  
			  return ret;
		  }
		  
		  public static String tirarAlfabeto(String str) {
			  
				//  System.out.println("TRATAR SUJEIRA :|"+str+"|");
				  
			        String ret = "";
			        
			        ret = str;
			        
			        
			        ret = ret.replace("a", "");
			     	ret = ret.replace("b", "");
			     	ret = ret.replace("b", "");
			     	ret = ret.replace("c", "");
			     	ret = ret.replace("d", "");
			     	ret = ret.replace("e", "");
			     	ret = ret.replace("f", "");
			     	ret = ret.replace("g", "");
			     	ret = ret.replace("h", "");
			     	ret = ret.replace("i", "");
			     	ret = ret.replace("j", "");
			     	ret = ret.replace("k", "");
			     	ret = ret.replace("l", "");
			     	ret = ret.replace("m", "");
			     	ret = ret.replace("n", "");
			     	ret = ret.replace("o", "");
			     	ret = ret.replace("p", "");
			     	ret = ret.replace("q", "");
			     	ret = ret.replace("r", "");
			     	ret = ret.replace("s", "");
			     	ret = ret.replace("t", "");
			     	ret = ret.replace("u", "");
			     	ret = ret.replace("w", "");
			     	ret = ret.replace("x", "");
			     	ret = ret.replace("y", "");
			     	ret = ret.replace("z", "");
			     	
			        ret = ret.replace("A", "");
			     	ret = ret.replace("B", "");
			     	ret = ret.replace("C", "");
			     	ret = ret.replace("D", "");
			     	ret = ret.replace("E", "");
			     	ret = ret.replace("F", "");
			     	ret = ret.replace("F", "");
			     	ret = ret.replace("G", "");
			     	ret = ret.replace("H", "");
			     	ret = ret.replace("I", "");
			     	ret = ret.replace("J", "");
			     	ret = ret.replace("K", "");
			     	ret = ret.replace("L", "");
			     	ret = ret.replace("M", "");
			     	ret = ret.replace("N", "");
			     	ret = ret.replace("O", "");
			     	ret = ret.replace("P", "");
			     	ret = ret.replace("Q", "");
			     	ret = ret.replace("R", "");
			     	ret = ret.replace("S", "");
			     	ret = ret.replace("T", "");
			     	ret = ret.replace("U", "");
			     	ret = ret.replace("W", "");
			     	ret = ret.replace("X", "");
			     	ret = ret.replace("Y", "");
			     	ret = ret.replace("Z", "");
			     	

		       		return ret;
			    }
		  
		  
		  public static String trataSujeira(String str) {
			  
			//  System.out.println("TRATAR SUJEIRA :|"+str+"|");
			  
		        String ret = "";
		        
		        ret = str;
		        
		        
		        ret = ret.replaceAll("(", "");
		        ret = ret.replaceAll(")", "");


	       		return ret;
		    }
		  
		  
		  private static boolean campoNumerico(String campo){           
		        return campo.matches("[0-9]+");   
		}
				  
		  		  
		  public static String toTitledCase(String nome){
			  
			  //System.out.println("STR ENTRADA= "+ nome);
			  
			  nome = " "+nome; 
			  	
			  String aux =""; // só é utilizada para facilitar 

		        try{ //Bloco try-catch utilizado pois leitura de string gera a exceção abaixo
		            for(int i = 0; i < nome.length(); ++i){
		                if( nome.substring(i, i+1).equals(" ") || nome.substring(i, i+1).equals("  "))
		                {
		                    aux += nome.substring(i+1, i+2).toUpperCase();
		                   // System.out.println("1= "+ aux);
		                }
		                else
		                {
		                    aux += nome.substring(i+1, i+2).toLowerCase();
		                    //System.out.println("2= "+ aux);
		                }
		        }
		        }catch(IndexOutOfBoundsException indexOutOfBoundsException){
		            //não faça nada. só pare tudo e saia do bloco de instrução try-catch
		        }
		        nome = aux;
		       // System.err.println(nome);

			  return nome;
			}  
		  
		  
		  public static String diretorio(String caminho) {
			
			  File diretorio = new File(caminho); // ajfilho é uma pasta!
			  if (!diretorio.exists()) {
			     diretorio.mkdir(); //mkdir() cria somente um diretório, mkdirs() cria diretórios e subdiretórios.
			  } else {
			     System.out.println("Diretório já existente");
			  }

			  
			  return caminho;
			  
		  }
		  
		  public static String lerExcel(String str) throws IOException{

			  
			  
					int contAux = 0;// Controle para pegar o Numer da Conta Juridica
					int contadorPosicao = 11; //Para saber a linha que passou
					
					//Conta Juridica Existente:
					int posicaoExiste = 0; // Pega Guardar a posicao da Conta Existente
					int posicaoExisteUnica = 0; // Pega Guardar a posicao da Conta Existente
					
					boolean cjExiste = false;
					boolean cjExisteUnica = true;
					boolean cjNova = false;
					boolean barrarGravacaoContaJaExiste = false;
					
					long contaJudicial = 0;
					String numeroDoProcesso = "";
					String numeroDoProcessoAuxi = "";
					String nomeDoReu = "";
					String numeroParcela = "";
					boolean temValorParcela = false;
					String posicaoEvalor = "";
					
					String E = "E1";  // NUMERO PROCESSO COLUNA CORRESPONDENTE DO EXCEL
					String B = "B1"; // NOME DO AUTOR COLUNA CORRESPONDENTE DO EXCEL
					//String N = "N1";
					
					ArrayList<String> arrayConteudoExiste = new ArrayList<String> ();
					ArrayList<String> arrayConteudoNovoComValor = new ArrayList<String> ();
					ArrayList<String> arrayConteudoNovoSemValor = new ArrayList<String> ();
					ArrayList<String> arrayConteudoCJExisteNaoUnica = new ArrayList<String> ();
					
					ArrayList<String> arrayConteudoCJExisteNaoUnicaIncluirValor = new ArrayList<String> ();
					String ConteudoCJExisteNaoUnicaIncluir = "";
					
					ArrayList<String> arrayConteudoCJExisteNaoUnicaIncluirParcela = new ArrayList<String> ();
					String ConteudoExistenumeroParcela = "";
					String parcela = "";
					String CJXLX = "";
					String CJPDF = "";
			 
					try {

						ZipSecureFile.setMinInflateRatio(-1.0d);
			            XSSFWorkbook workbook = new XSSFWorkbook(excelBB);
			            XSSFSheet sheet = workbook.getSheetAt(0);
			            Row row = sheet.getRow(0);  
			            
			            getSetCEF.setCjExisteUnica(true);
			            getSetCEF.setCjExiste(false);
			            getSetCEF.setCjNova(true);
			            
			            CellReference cellReferencePAR = null;
			            Row rowLPAR = null;
			            Cell cellLPAR = null;
			            
			            CellReference cellReferenceVAL = null;
			            Row rowLVAL = null;
			            Cell cellLVAL = null;		           
			            
			            CellReference cellReferenceCJ = null;
			            Row rowLCJ = null;
			            Cell cellLCJ = null;
			            
			            
	
			            	for (int i = 11; i < 10000; i++) //Comeca do 11 para inicinar na linha 11
						{
			            	      try 
			            	      {
			            	    	  	//LENDO AS COLUNAS N1, N2 , N3  (VALOR ATUALIZADO)
//				            		 N = "N"+i;
//				            		 cellReferencePAR = new CellReference(N);   //Ferencia Coluna M usado na Conta Judicial
//				            		 rowLPAR = sheet.getRow(cellReferencePAR.getRow());	 //Ferencia Linha usado na Conta Judicial
//				            	     cellLPAR = rowLPAR.getCell(cellReferencePAR.getCol());
	
			            	    	  
			            	    	  
			            	    	  	//LENDO AS COLUNAS B // NOME AUTOR  COLUNA CORRESPONDENTE DO EXCEL
				            		 B = "L"+i;
				            		 cellReferenceVAL = new CellReference(B);  
				            		 rowLVAL = sheet.getRow(cellReferenceVAL.getRow());	
				            	     cellLVAL = rowLVAL.getCell(cellReferenceVAL.getCol());
	
			            	    	  
			            	    	  
				            	     //LENDO AS COLUNAS E // NUMERO PROCESSO  COLUNA CORRESPONDENTE DO EXCEL
				            		 E = "E"+i;			            		 
				            		 cellReferenceCJ = new CellReference(E);   
				            		 rowLCJ = sheet.getRow(cellReferenceCJ.getRow());	 
				            	     contaJudicial = 0;
		
			            	    	     cellLCJ = rowLCJ.getCell(cellReferenceCJ.getCol());  
				            	     

				            	     
			            	    	 
			            	    	     	if(cellLCJ.CELL_TYPE_NUMERIC==cellLCJ.getCellType()) {
			            	    	     		numeroDoProcesso = String.valueOf(cellLCJ.getNumericCellValue()) ;
			            	    	     	}
			            	    	     	if(cellLCJ.CELL_TYPE_STRING==cellLCJ.getCellType()) {
			            	    	     		numeroDoProcesso =	cellLCJ.getStringCellValue();
			            	    	     	}
			            	    	     
			            	    	     	
			            	    	     	numeroDoProcesso 		= numeroDoProcesso.replaceAll("[^0-9]","");// processo do excel
			            	    	     	numeroDoProcesso 		= completeToLeft(numeroDoProcesso, '0', 20);
			            	    	     	
			            	    	     	numeroDoProcessoAuxi 	= getSetCEF.getNumeroProcessoAuxiliar();//processo do PDF
			            	    	     	numeroDoProcessoAuxi		= completeToLeft(numeroDoProcessoAuxi, '0', 20);
			            	    	     	
			            	    	     	
			            	    	     	
			            	    	     	
//			            	    	     	System.out.println(" numeroDoProcessoAuxi " + numeroDoProcesso);
//			            	    	     	System.out.println(" numeroDoProcesso.substring(0, 9) :  " + numeroDoProcesso.substring(0, 9));
//			            	    	     	System.out.println(" numeroDoProcesso.substring(7, 13) :  " + numeroDoProcesso.substring(7, 13));
//			            	    	     	System.out.println(" numeroDoProcesso.substring(0, 7) :  " + numeroDoProcesso.substring(0, 7));
//			            	    	     	System.out.println(" numeroDoProcesso.substring(9, 13) :  " + numeroDoProcesso.substring(9, 13));
			            	    	     	
			            	    	     	
			            	    	     	
			            	    	     	if( numeroDoProcesso.equals(numeroDoProcessoAuxi) ||
			            	    	     		numeroDoProcesso.substring(0, 9).equals(numeroDoProcessoAuxi.substring(0,9))	||
			            	    	     		numeroDoProcesso.substring(7, 13).equals(numeroDoProcessoAuxi.substring(7,13))	||
			            	    	     		numeroDoProcesso.substring(0, 7).equals(numeroDoProcessoAuxi.substring(0,7)) && numeroDoProcesso.substring(9, 13).equals(numeroDoProcessoAuxi.substring(9,13)) 
			            	    	     	  ) 
			                        	{
				            	    	  		getSetCEF.setPosicaoExiste(i);	
			                        	 			                        	 	
		                        	 		//AQUI O NOME DO REU
			   			            	     if(cellLVAL.CELL_TYPE_NUMERIC==cellLVAL.getCellType()) {
			   			            	    	 	nomeDoReu = String.valueOf(cellLVAL.getNumericCellValue()) ;
						            	     }
			   			            	     if(cellLVAL.CELL_TYPE_STRING==cellLVAL.getCellType()) {
			   			            	    	 	nomeDoReu =	cellLVAL.getStringCellValue();
						            	     }
			   			            	     
			   			            	     
			   			            	     //AQUI CHAMAR PARA PREENCHER A LINHA DO EXCEL	
			   			            	     EscreverContaNovaComValor(); 
			   			            	  
			   			            	     
							            	//PAUSA PARA PODER GRAVAR O EXCEL
							            	try {  
									        Thread.sleep( 100 );  
									     } 
						  			     catch (InterruptedException e) {  
									         e.printStackTrace();  
									     } 
							            	
							            	//PARA O FOUR GERAL E VAI PARA O PROXIMO
							            	break;
		                        	 		
			                        	}
			            	    	     	//elseif(INICIAR COMPARACAO DE QUEBRA DE PROCESSO DEPOIS NOME)
	
							  }catch(Exception e){
								//  contadorPosicao ++;	
								  break;
						      }
						}
	
	
			            


		        } catch (IOException e) {
		            e.printStackTrace();
		        }				
				return "";
		  }
		  
		  
		  public static String completeToLeft(String value, char c, int size) {
				String result = value;
				while (result.length() < size) {
					result = c + result;
				}
				return result;
			}
		  
		  
		  public static void EscreverContaNovaComValor() throws IOException {
			  try{
				  
				  	XSSFWorkbook a = null; 
				  	
			         a = new XSSFWorkbook(new FileInputStream(fo));
			        
			         XSSFSheet my_sheet = null;
			         
			         my_sheet = a.getSheetAt(0);
			        
			        
			        System.out.println("3 -  EscreverContaNovaComValor GRAVAR NA LINHA :  " + getSetCEF.getPosicaoExiste());
			        
			        
			        //Centro Azul Claro
			        XSSFCellStyle style2 = a.createCellStyle();
			        style2.setAlignment ( XSSFCellStyle.ALIGN_CENTER ) ; 

			        
			        
			        //Centro Azul Claro
			        XSSFCellStyle style3 = a.createCellStyle();
			        style3.setFillForegroundColor(new XSSFColor(new java.awt.Color(89, 179, 8)));
			        style3.setAlignment ( XSSFCellStyle.ALIGN_CENTER ) ; 
			        style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
			        style3.setBorderBottom(CellStyle.BORDER_THIN);
			        style3.setBottomBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        style3.setBorderLeft(CellStyle.BORDER_THIN);
			        style3.setLeftBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        style3.setBorderRight(CellStyle.BORDER_THIN);
			        style3.setRightBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        style3.setBorderTop(CellStyle.BORDER_THIN);
			        style3.setTopBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        
			        


			       // my_sheet.createRow(getSetCEF.getPosicaoExiste()-1);
			        
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).createCell(1);
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(1).setCellValue(getSetCEF.getAutor());
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(1).setCellStyle(style2);
			        
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).createCell(2);
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(2).setCellValue(getSetCEF.getReu());
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(2).setCellStyle(style2);
		        
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(3);
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(3).setCellValue(getSetCEF.getCNPJ());
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(3).setCellStyle(style2);
			        
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).createCell(4);
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(4).setCellValue(getSetCEF.getNumeroProcesso());
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(4).setCellStyle(style3);
			        
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).createCell(5);
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(5).setCellValue(getSetCEF.getVara());
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(5).setCellStyle(style2);
			        
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).createCell(6);
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(6).setCellValue(getSetCEF.getComarca());
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(6).setCellStyle(style2);
			        
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(7);
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(7).setCellValue(getSetCEF.getEstado());
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(7).setCellStyle(style2);
//
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(8);
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(8).setCellValue("Trabalhista");
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(8).setCellStyle(style2);
//			        
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(9);
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(9).setCellValue(getSetCEF.getDataDeposito());
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(9).setCellStyle(data);
//
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(10);
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(10).setCellValue(getSetCEF.getValorOriginal());
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(10).setCellStyle(style1);
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(10).setCellType(XSSFCell.CELL_TYPE_STRING);
//			        
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(11);
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(11).setCellValue(getSetCEF.getValorAtualizado());			        
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(11).setCellStyle(style1);
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(11).setCellType(XSSFCell.CELL_TYPE_STRING);
//			        
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(12);
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(12).setCellValue(getSetCEF.getContaJuridica());
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(12).setCellStyle(style3);
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(12).setCellType(XSSFCell.CELL_TYPE_STRING);
			        
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(13);
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(13).setCellValue(getSetCEF.getParcela());			        
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(13).setCellStyle(style2);
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(13).setCellType(XSSFCell.CELL_TYPE_STRING);

			        FileOutputStream outputStream  = null;
			        outputStream = new FileOutputStream(new File(excelBB));
			        
			        a.write(outputStream);
			        outputStream.close();//Close in finally if possible
			        outputStream = null;
			        
		        }catch(Exception e){
		        		System.out.println("3 -  EscreverContaNovaComValor GRAVAR NA LINHA EROOO");
		        		System.out.println(" ERRO: " + e);
		        }
			}	
		  		 
		  
	  
		  
}