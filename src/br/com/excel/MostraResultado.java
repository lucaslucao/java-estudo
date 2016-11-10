package br.com.excel;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JPanel;
import javax.swing.JTextField;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MostraResultado extends JFrame {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private JTextField fieldConcurso;
	private JTextField fieldPrimeirad;
	private JTextField fieldSegundad;
	private JTextField fieldTerceirad;
	private JTextField fieldQuartad;
	private JTextField fieldQuintad;
	private JTextField fieldSextad;

	public MostraResultado() {
		setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		setTitle("Sorteia");
		getContentPane().setLayout(null);
		JLabel lblConcurso = new JLabel("Concurso: ");
		lblConcurso.setBounds(36, 36, 68, 14);
		getContentPane().add(lblConcurso);

		fieldConcurso = new JTextField();
		fieldConcurso.setBounds(98, 33, 86, 20);
		getContentPane().add(fieldConcurso);
		fieldConcurso.setColumns(10);

		fieldPrimeirad = new JTextField();
		fieldPrimeirad.setBounds(31, 79, 46, 20);
		getContentPane().add(fieldPrimeirad);
		fieldPrimeirad.setColumns(10);

		fieldSegundad = new JTextField();
		fieldSegundad.setBounds(93, 79, 46, 20);
		getContentPane().add(fieldSegundad);
		fieldSegundad.setColumns(10);

		fieldTerceirad = new JTextField();
		fieldTerceirad.setBounds(155, 79, 46, 20);
		getContentPane().add(fieldTerceirad);
		fieldTerceirad.setColumns(10);

		fieldQuartad = new JTextField();
		fieldQuartad.setBounds(211, 79, 46, 20);
		getContentPane().add(fieldQuartad);
		fieldQuartad.setColumns(10);

		fieldQuintad = new JTextField();
		fieldQuintad.setBounds(273, 79, 46, 20);
		getContentPane().add(fieldQuintad);
		fieldQuintad.setColumns(10);

		fieldSextad = new JTextField();
		fieldSextad.setBounds(329, 79, 46, 20);
		getContentPane().add(fieldSextad);
		fieldSextad.setColumns(10);

		JButton btnPesquisar = new JButton("Pesquisar");
		btnPesquisar.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				try {
					sorteia();
				} catch (Exception e) {
					// TODO: handle exception
				}
			}
		});
		btnPesquisar.setBounds(216, 32, 141, 23);
		getContentPane().add(btnPesquisar);

		JPanel panel = new JPanel();
		panel.setBounds(31, 111, 344, 57);
		getContentPane().add(panel);

		JList list = new JList();
		panel.add(list);

		JButton btnVerificar = new JButton("Verificar");
		btnVerificar.setBounds(240, 181, 98, 26);
		getContentPane().add(btnVerificar);

		JLabel lblExiste = new JLabel("");
		lblExiste.setBounds(58, 180, 143, 26);
		getContentPane().add(lblExiste);

		setVisible(true);
		setSize(420, 300);
	}

	public void sorteia() throws Exception {
		File file = new File("megasena.xlsx");
		FileInputStream fIP = new FileInputStream(file);
		// Get the workbook instance for XLSX file
		XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		if (file.isFile() && file.exists()) {
			System.out.println("Lendo arquivo..."/* "openworkbook.xlsx file open successfully." */);
		} else {
			System.out.println("Error to open openworkbook.xlsx file.");
		}

		int[][] matriz = percorre(workbook);

		int concurso = Integer.parseInt(fieldConcurso.getText());
		for (int i = 0; i < matriz.length; i++) {
			if (matriz[i][0] == concurso) {
				for (int j = 0; j < matriz[i].length; j++) {
					System.out.print(matriz[i][j] + " ");

				}
				fieldPrimeirad.setText(Integer.toString(matriz[i][1]));
				fieldSegundad.setText(Integer.toString(matriz[i][2]));
				fieldTerceirad.setText(Integer.toString(matriz[i][3]));
				fieldQuartad.setText(Integer.toString(matriz[i][4]));
				fieldQuintad.setText(Integer.toString(matriz[i][5]));
				fieldSextad.setText(Integer.toString(matriz[i][6]));
				break;
			}
			// System.out.println();
		}

	}

	private int[][] percorre(XSSFWorkbook workbook) {
		XSSFSheet sheet = workbook.getSheetAt(0);
		XSSFRow row;
		XSSFCell cell;
		Iterator rows = sheet.rowIterator();
		int matriz[][] = new int[sheet.getLastRowNum() + 1][7];
		System.out.println();

		int lin = 0;

		while (rows.hasNext()) {
			row = (XSSFRow) rows.next();
			Iterator cells = row.cellIterator();
			int col = 0;
			while (cells.hasNext()) {
				cell = (XSSFCell) cells.next();
				if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
					System.out.print(cell.getStringCellValue() + " ");
				} else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {

					Double numericCellValue = cell.getNumericCellValue();
					/* Variavel tipo classe Double com D maiusculo */
					int valor = numericCellValue.intValue();
					// System.out.print(valor + " ");
					matriz[lin][col] = valor;
				} else {
					// throw new Exception("Erro");
				}
				col++;
			}
			lin++;
			// System.out.println();
		}
		return matriz;
	}

	public static void main(String args[]) throws Exception {

		new MostraResultado();
	}
}
