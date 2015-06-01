package excel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Locale;
import java.util.Scanner;

import jxl.CellReferenceHelper;
import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.biff.DisplayFormat;
import jxl.format.CellFormat;
import jxl.format.UnderlineStyle;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.NumberFormat;
import jxl.write.NumberFormats;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class WriteExcel {
	private static int PROCESSORS_COLUMN = 0;
	private static int PARALLEL_RUNTIME_COLUMN = 1;
	private static int SPEEDUP_COLUMN = 2;
	private static int SEQUENTIAL_PROBLEM_SIZE_COLUMN = 4;
	private static int SEQUENTIAL_RUNTIME_COLUMN = 5;
	private static String fileName = "file.xls";
	private static double[] seqRuntimes = {10.0,12.0};
	private static double[][] runtimes = { { 7,8 }, { 9, 10 } };
	private static double[] sequentialRuntimes = {5,6};
	private static int[] problemSizes;
	private static int[] processors;
	
	public enum DataType {
	    INT, DOUBLE
	}

	public static void main(String[] args) throws WriteException, IOException {
		try {
			// Write comma separated runtimes into 2D-Array
			Scanner sc = new Scanner(new File("parallel_runtimes.txt"));
			int index = 0;
			while (sc.hasNextLine()) {
				String line = sc.nextLine();
				String[] separatedValues = line.split(",");
				for (int i=0;i<separatedValues.length;i++) {
					runtimes[index][i] = Double.parseDouble(separatedValues[i]);
				}
				index++;
			}
			sc.close();
			sc = new Scanner(new File("sequential_runtimes.txt"));
			String line = sc.nextLine();
			String[] separatedValues = line.split(",");
			for (int i=0;i<separatedValues.length;i++) {
				sequentialRuntimes[i] = Double.parseDouble(separatedValues[i]);
			}
			processors = parseIntegers("processors.txt");
			problemSizes = parseIntegers("problem_sizes.txt");

			
			WritableWorkbook workbook = Workbook.createWorkbook(new File(
					fileName));
			WritableSheet writableSheet = workbook.createSheet("Sheet1", 0);
			WritableCellFormat floatFormat = new WritableCellFormat (NumberFormats.FLOAT);
			WritableCellFormat integerFormat = new WritableCellFormat (NumberFormats.INTEGER);
			
			CellView integerCellView = new CellView();
		    integerCellView.setAutosize(true);
		    integerCellView.setFormat(integerFormat);
		    writableSheet.setColumnView(PROCESSORS_COLUMN, integerCellView);
		    writableSheet.setColumnView(SEQUENTIAL_PROBLEM_SIZE_COLUMN, integerCellView);
		    
		    CellView floatCellView = new CellView();
		    floatCellView.setAutosize(true);
		    floatCellView.setFormat(floatFormat);
		    writableSheet.setColumnView(PARALLEL_RUNTIME_COLUMN, floatCellView);
		    writableSheet.setColumnView(SPEEDUP_COLUMN, floatCellView);
		    writableSheet.setColumnView(SEQUENTIAL_RUNTIME_COLUMN, floatCellView);

			// sequential
			Label seqProblemSizeLabel = new Label(SEQUENTIAL_PROBLEM_SIZE_COLUMN, 0, "problem size");
			Label seqRuntimeLabel = new Label(SEQUENTIAL_RUNTIME_COLUMN, 0, "runtime (s)");
			writableSheet.addCell(seqProblemSizeLabel);
			writableSheet.addCell(seqRuntimeLabel);
			for (int i = 0; i < seqRuntimes.length; i++) {
				Number problemSize = new Number(SEQUENTIAL_PROBLEM_SIZE_COLUMN, i+1, problemSizes[i]);
				Number runtime = new Number(SEQUENTIAL_RUNTIME_COLUMN, i+1, seqRuntimes[i]);
				writableSheet.addCell(problemSize);
				writableSheet.addCell(runtime);
			}
			
			// parallel
			Label labelProcessors = new Label(PROCESSORS_COLUMN, 0, "processors");
			Label labelRuntime = new Label(PARALLEL_RUNTIME_COLUMN, 0, "runtime (s)");
			Label labelSpeedup = new Label(SPEEDUP_COLUMN, 0, "speedup");
			writableSheet.addCell(labelProcessors);
			writableSheet.addCell(labelRuntime);
			writableSheet.addCell(labelSpeedup);
			int yPos = 1;
			int seqRuntimeRow = yPos;
			for (int i = 0; i < runtimes.length; i++) {
				StringBuffer absSeqRuntimeCellReference = new StringBuffer();
				CellReferenceHelper.getCellReference(SEQUENTIAL_RUNTIME_COLUMN,true,seqRuntimeRow,true,absSeqRuntimeCellReference);
				for (int j = 0; j < processors.length; j++) {
					Number processor = new Number(PROCESSORS_COLUMN, yPos, processors[j]);
					Number runtime = new Number(PARALLEL_RUNTIME_COLUMN, yPos, runtimes[i][j]);
					String parRuntimeCellReference = CellReferenceHelper.getCellReference(PARALLEL_RUNTIME_COLUMN, yPos);
					String speedupFormula = absSeqRuntimeCellReference + "/" + parRuntimeCellReference;
					Formula speedup = new Formula(SPEEDUP_COLUMN, yPos, speedupFormula);
					writableSheet.addCell(processor);
					writableSheet.addCell(runtime);
					writableSheet.addCell(speedup);
					yPos++;
				}
				seqRuntimeRow++;
			}
			workbook.write();
			workbook.close();
		} catch (WriteException e) {

		}
	}
	
	private static int[] parseIntegers(String fileName) throws FileNotFoundException {
		Scanner sc = new Scanner(new File(fileName));
		String line = sc.nextLine();
		String[] separatedValues = line.split(",");
		
		int[] result = new int[separatedValues.length];
		for (int i=0;i<separatedValues.length;i++) {
			result[i] = Integer.parseInt(separatedValues[i]);
		}
		sc.close();
		return result;
	}
}