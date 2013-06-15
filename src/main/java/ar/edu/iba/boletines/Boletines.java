package ar.edu.iba.boletines;

import java.awt.Color;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.InputStream;
import java.util.Arrays;
import java.util.Calendar;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JSpinner;
import javax.swing.JTextArea;
import javax.swing.SwingConstants;

import org.eclipse.birt.core.framework.Platform;
import org.eclipse.birt.report.engine.api.EngineConfig;
import org.eclipse.birt.report.engine.api.EngineConstants;
import org.eclipse.birt.report.engine.api.HTMLRenderOption;
import org.eclipse.birt.report.engine.api.IPDFRenderOption;
import org.eclipse.birt.report.engine.api.IRenderOption;
import org.eclipse.birt.report.engine.api.IReportEngine;
import org.eclipse.birt.report.engine.api.IReportEngineFactory;
import org.eclipse.birt.report.engine.api.IReportRunnable;
import org.eclipse.birt.report.engine.api.IRunAndRenderTask;
import org.eclipse.birt.report.engine.api.PDFRenderOption;
import org.eclipse.birt.report.engine.api.RenderOption;

public class Boletines {

	private static String[] arguments;
	private static String REPORT_NAME = "boletinIBA.rptdesign";
	private JFrame frmBoletinesIba;
	private JTextArea lblExcelFile;
	private JTextArea lblOutputPath;
	private JSpinner anio;
	private JButton btnGenerar;
	private JLabel lblGenerandoReporte;
	private final JFileChooser fc = new JFileChooser();
	private static Logger log = Logger.getLogger("ar.edu.iba.boletines.Boletines");
	private IReportEngine engine;
	private IReportRunnable design;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		arguments = args;
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Boletines window = new Boletines();
					window.frmBoletinesIba.setVisible(true);
				} catch (Exception e) {
					log.log(Level.SEVERE, "No se pudo inicializar la aplicaci—n", e);
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public Boletines() {
		log.info(Arrays.asList(arguments).toString());
		initialize();
		reportEngineInitialization();
	}

	private void reportEngineInitialization() {
		try {
			final EngineConfig config = new EngineConfig();
			Platform.startup(config);
			final IReportEngineFactory FACTORY = (IReportEngineFactory) Platform
					.createFactoryObject(IReportEngineFactory.EXTENSION_REPORT_ENGINE_FACTORY);
			engine = FACTORY.createReportEngine(config);
			engine.changeLogLevel(Level.WARNING);
			InputStream inputStream = getClass().getResourceAsStream("/" + REPORT_NAME);
			// Open the report design
			design = engine.openReportDesign(inputStream);
		} catch (Exception e) {
			log.log(Level.SEVERE, "Error iniciando motor de reportes", e);
		}
	}

	private void selectExcelFile(ActionEvent event) {
		File excelFile = getFileFromPicker(false);
		lblExcelFile.setText(excelFile.getAbsolutePath());
	}

	private void selectOutpuFile(ActionEvent event) {
		File outputFile = getFileFromPicker(true);
		lblOutputPath.setText(outputFile.getAbsolutePath());
	}

	private File getFileFromPicker(boolean onlyDirectories) {
		File file = null;
		fc.setFileSelectionMode(onlyDirectories ? JFileChooser.DIRECTORIES_ONLY
				: JFileChooser.FILES_AND_DIRECTORIES);
		int returnVal = fc.showOpenDialog(frmBoletinesIba);
		if (returnVal == JFileChooser.APPROVE_OPTION) {
			file = fc.getSelectedFile();
		}
		return file;
	}

	private void generate(ActionEvent event) {
		btnGenerar.setVisible(false);
		lblGenerandoReporte.setVisible(true);

		EventQueue.invokeLater(new Runnable() {
			public void run() {
				generateReport();
			}
		});
	}

	private void generateReport() {
		try {
			// Create task to run and render the report,
			final IRunAndRenderTask task = engine.createRunAndRenderTask(design);
			// Set parent classloader for engine
			task.getAppContext().put(EngineConstants.APPCONTEXT_CLASSLOADER_KEY,
					Boletines.class.getClassLoader());

			final IRenderOption options = new RenderOption();
			options.setOutputFormat("PDF");
			String outputFileName = lblOutputPath.getText() + File.separator
					+ lblExcelFile.getText().replaceAll("^.*\\" + File.separator + "|\\.xlsx|\\.xls", "")
					+ ".pdf";
			options.setOutputFileName(outputFileName);
			task.setParameterValue("ANIO", anio.getValue());
			task.setParameterValue("XLS_PATH", lblExcelFile.getText());
			log.info("Archivo de entrada: " + lblExcelFile.getText());
			log.info("Archivo de salida: " + outputFileName);
			if (options.getOutputFormat().equalsIgnoreCase("html")) {
				final HTMLRenderOption htmlOptions = new HTMLRenderOption(options);
				htmlOptions.setImageDirectory("img");
				htmlOptions.setHtmlPagination(false);
				htmlOptions.setHtmlRtLFlag(false);
				htmlOptions.setEmbeddable(false);
				htmlOptions.setSupportedImageFormats("PNG");

				// set this if you want your image source url to be altered
				// If using the setBaseImageURL, make sure to set image handler
				// to HTMLServerImageHandler
				// htmlOptions.setBaseImageURL("http://myhost/prependme?image=");
			} else if (options.getOutputFormat().equalsIgnoreCase("pdf")) {
				final PDFRenderOption pdfOptions = new PDFRenderOption(options);
				pdfOptions.setOption(IPDFRenderOption.PAGE_OVERFLOW, IPDFRenderOption.FIT_TO_PAGE_SIZE);
				pdfOptions.setOption(IPDFRenderOption.PAGE_OVERFLOW,
						IPDFRenderOption.OUTPUT_TO_MULTIPLE_PAGES);
			}

			task.setRenderOption(options);

			// run and render report
			task.run();

			task.close();
		} catch (Exception ex) {
			log.log(Level.SEVERE, "Error generando reporte", ex);
		} finally {
			lblGenerandoReporte.setVisible(false);
			btnGenerar.setVisible(true);
		}
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frmBoletinesIba = new JFrame();
		frmBoletinesIba.setTitle("Boletines - IBA");
		frmBoletinesIba.setBounds(100, 100, 550, 220);
		frmBoletinesIba.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frmBoletinesIba.getContentPane().setLayout(null);

		JLabel lblCicloLectivo = new JLabel("Ciclo lectivo:");
		lblCicloLectivo.setBounds(6, 12, 100, 16);
		frmBoletinesIba.getContentPane().add(lblCicloLectivo);

		anio = new JSpinner();
		anio.setBounds(135, 6, 80, 28);
		anio.setValue(Calendar.getInstance().get(Calendar.YEAR));
		frmBoletinesIba.getContentPane().add(anio);

		JButton btnArchivoExcel = new JButton("Archivo Excel");
		btnArchivoExcel.setBounds(6, 50, 120, 29);
		btnArchivoExcel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent event) {
				selectExcelFile(event);
			}
		});
		frmBoletinesIba.getContentPane().add(btnArchivoExcel);

		JButton btnDirSalida = new JButton("Dir. Salida");
		btnDirSalida.setBounds(6, 100, 120, 29);
		btnDirSalida.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent event) {
				selectOutpuFile(event);
			}
		});
		frmBoletinesIba.getContentPane().add(btnDirSalida);

		lblExcelFile = new JMultilineLabel("");
		lblExcelFile.setBounds(135, 40, 410, 50);
		frmBoletinesIba.getContentPane().add(lblExcelFile);

		lblOutputPath = new JMultilineLabel(System.getProperty("user.home"));
		lblOutputPath.setBounds(135, 90, 410, 50);
		frmBoletinesIba.getContentPane().add(lblOutputPath);

		btnGenerar = new JButton("Generar");
		btnGenerar.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent event) {
				generate(event);
			}
		});
		btnGenerar.setBounds(216, 143, 117, 29);
		frmBoletinesIba.getContentPane().add(btnGenerar);

		lblGenerandoReporte = new JLabel("Generando Boletines");
		lblGenerandoReporte.setToolTipText("Aguarde mientras se generan los boletines");
		lblGenerandoReporte.setVisible(false);
		lblGenerandoReporte.setFont(new Font("Lucida Grande", Font.BOLD, 14));
		lblGenerandoReporte.setHorizontalAlignment(SwingConstants.CENTER);
		lblGenerandoReporte.setForeground(new Color(0, 153, 0));
		lblGenerandoReporte.setBounds(181, 148, 187, 16);
		frmBoletinesIba.getContentPane().add(lblGenerandoReporte);
	}

	public class JMultilineLabel extends JTextArea {
		private static final long serialVersionUID = 1L;

		public JMultilineLabel(String text) {
			super(text);
			setEditable(false);
			setCursor(null);
			setOpaque(false);
			setFocusable(false);
			setWrapStyleWord(true);
			setLineWrap(true);
		}
	}
}
