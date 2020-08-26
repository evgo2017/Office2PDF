package test;
import java.io.File;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.*;

public class test {
	private static final int wdFormatPDF = 17;
	private static final int xlTypePDF = 0;
	private static final int ppSaveAsPDF = 32;
 
	public static void main(String[] args) {
		//放置需转换的文件的 file 文件夹目录
		String filepath = "C://test";
		 // 获得指定文件对象  
        File file = new File(filepath);   
        // 获得该文件夹内的所有文件   
        File[] fileList = file.listFiles();   
        System.out.println("开始\n====================");
        long start = System.currentTimeMillis();
		for(int i = 0; i<fileList.length; i++) {
			test.convert2PDF(fileList[i].toString(), changeFileSufix2PDF(fileList[i].toString()));
		}
		long end = System.currentTimeMillis();
		System.out.println("\n====================\n全部转换完成..用时：" + (end - start) + "ms.");
		System.out.println("====================\n结束");
	}
	public static boolean convert2PDF(String inputFile, String pdfFile) {
		// 获得文件
		String suffix = getFileSufix(inputFile);
		File file = new File(inputFile);
		if (!file.exists()) {
			return false;
		}
		if (suffix.equals("pdf")) {
			return false;
		}
		if (suffix.equals("doc") || suffix.equals("docx")
				|| suffix.equals("txt")) {
			return word2PDF(inputFile, pdfFile);
		} else if (suffix.equals("ppt") || suffix.equals("pptx")) {
			return ppt2PDF(inputFile, pdfFile);
		} else if (suffix.equals("xls") || suffix.equals("xlsx")) {
			return excel2PDF(inputFile, pdfFile);
		} else {
			return false;
		}
	}
	public static String getFileSufix(String fileName) {
		int splitIndex = fileName.lastIndexOf(".");
		return fileName.substring(splitIndex + 1);
	}
	public static String changeFileSufix2PDF(String fileName) {
		int splitIndex = fileName.lastIndexOf(".");
		return fileName.substring(0, splitIndex) + ".pdf";
	}
	// word转换为pdf
	public static boolean word2PDF(String inputFile, String pdfFile) {
		System.out.println("\n====================\n启动Word..."); 
		long start = System.currentTimeMillis(); 
		ActiveXComponent app = null;
		Dispatch docs = null;
		Dispatch doc = null;
		try {
			// 打开word应用程序
			app = new ActiveXComponent("Word.Application");
			// 设置word不可见
			app.setProperty("Visible", false);
			// 获得word中所有打开的文档,返回Documents对象
			System.out.println("打开文档..." + inputFile);
			docs = app.getProperty("Documents").toDispatch();
			// 调用Documents对象中Open方法打开文档，并返回打开的文档对象Document
			doc = Dispatch.call(docs, "Open", inputFile, false, true)
					.toDispatch();
			// 调用Document对象的SaveAs方法，将文档保存为pdf格式
			/*
			 * Dispatch.call(doc, "SaveAs", pdfFile, wdFormatPDF
			 * //word保存为pdf格式宏，值为17 );
			 */
			// word保存为pdf格式宏，值为17
			System.out.println("转换文档到PDF..." + pdfFile);
			Dispatch.call(doc, "ExportAsFixedFormat", pdfFile, wdFormatPDF);
			// 关闭文档
			Dispatch.call(doc, "Close", false);
			// 关闭word应用程序
			app.invoke("Quit", 0);
			
			long end = System.currentTimeMillis();
			System.out.println("转换完成..用时：" + (end - start) + "ms.\n====================");
			
			return true;
		} catch (Exception e) {
			System.out.println("-----------Error:文档转换失败：" + e.getMessage() +"\n====================");
			return false;
		} finally {
			ComThread.Release();
		}
	}
	// excel转换为pdf
	public static boolean excel2PDF(String inputFile, String pdfFile) {
		System.out.println("\n====================\n启动Excel..."); 
		long start = System.currentTimeMillis(); 
		ActiveXComponent app = null;
		Dispatch docs = null;
		Dispatch doc = null;
		try {
			app = new ActiveXComponent("Excel.Application");
			app.setProperty("Visible", false);
			System.out.println("打开文档..." + inputFile);
			Dispatch excels = app.getProperty("Workbooks").toDispatch();
			Dispatch excel = Dispatch.call(excels, "Open", inputFile, false,
					true).toDispatch();
			System.out.println("转换文档到PDF..." + pdfFile);
			Dispatch.call(excel, "ExportAsFixedFormat", xlTypePDF, pdfFile);
			
			Dispatch.call(excel, "Close", false);
			app.invoke("Quit");
			
			long end = System.currentTimeMillis();
			System.out.println("转换完成..用时：" + (end - start) + "ms.\n====================");
			
			return true;
		} catch (Exception e) {
			System.out.println("========Error:文档转换失败：=========");
			System.out.println("(若找不到要打印的任何内容，可能是因为文件无内容)");
			System.out.println(e.getMessage()+"\n====================");
			return false;
		} finally {
			ComThread.Release();
		}
	}
	// ppt转换为pdf
	public static boolean ppt2PDF(String inputFile, String pdfFile) {
		System.out.println("\n====================\n启动PPT..."); 
		long start = System.currentTimeMillis(); 
		ActiveXComponent app = null;
		Dispatch ppts = null;
		Dispatch ppt = null;
		try {
			app = new ActiveXComponent("PowerPoint.Application");
			// app.setProperty("Visible", msofalse);
			System.out.println("打开文档..." + inputFile);
			ppts = app.getProperty("Presentations").toDispatch();
			
			ppt = Dispatch.call(ppts, "Open", inputFile, true,// ReadOnly
					true,// Untitled指定文件是否有标题
					false// WithWindow指定文件是否可见
					).toDispatch();
			System.out.println("转换文档到PDF..." + pdfFile);
			Dispatch.call(ppt, "SaveAs", pdfFile, ppSaveAsPDF);
			
			Dispatch.call(ppt, "Close");
			
			long end = System.currentTimeMillis();
			System.out.println("转换完成..用时：" + (end - start) + "ms.");
			app.invoke("Quit");
			return true;
		} catch (Exception e) {	
			System.out.println("========Error:文档转换失败：=========");
			System.out.println("若发生PowerPoint 存储此文件时发生错误，可能是文件无内容");
			System.out.println(e.getMessage()+"\n====================");
			return false;
		} finally {
			ComThread.Release();
		}
	}
}