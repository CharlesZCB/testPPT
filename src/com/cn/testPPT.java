package com.cn;


import java.awt.Color;

import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.RenderingHints;

import java.awt.geom.Rectangle2D;

import java.awt.image.BufferedImage;

import java.io.File;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;
import java.util.UUID;

import org.apache.poi.xslf.usermodel.XMLSlideShow;

import org.apache.poi.xslf.usermodel.XSLFShape;

import org.apache.poi.xslf.usermodel.XSLFSlide;

import org.apache.poi.xslf.usermodel.XSLFTextParagraph;

import org.apache.poi.xslf.usermodel.XSLFTextRun;

import org.apache.poi.xslf.usermodel.XSLFTextShape;


 

public class testPPT {
public static final int ppSaveAsPNG = 17;
	
	public static void main(String[] args) {
		//ʵ�ʾ��Ǵ�����һ�� �ļ� ��Ҫ����ļ���Orignalname ����ֵ
		String infoString = doPPTtoImage2007(new File("C:\\Users\\ThinkPad\\Desktop\\Test.ppt"), "C:\\Users\\ThinkPad\\Desktop\\img", "a", "jpg");
		System.err.println(infoString);
	}

	/**
	 * ����PPTX���� 
	 * @param slide
	 */
    private static void setFont(XSLFSlide slide) {
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTextShape) {
            	XSLFTextShape txtshape = (XSLFTextShape)shape ;
                for (XSLFTextParagraph paragraph : txtshape.getTextParagraphs()) {
                    List<XSLFTextRun> truns = paragraph.getTextRuns();
                    for (XSLFTextRun trun : truns) {
                    	trun.setFontFamily("����");
                    }
                }
            }
        }
    }
	

	public static String doPPTtoImage2007(File file,String path,String picName,String picType) {

		try {

			boolean isppt = checkFile(file);

			if (!isppt) {
				System.out.println("�ļ���ʽֻ����PPT ���� PPTX");
			return "format error";
			}

			FileInputStream is = new FileInputStream(file);

			XMLSlideShow xmlSlideShow = new XMLSlideShow(is);

			XSLFSlide[] xslfSlides = xmlSlideShow.getSlides();

			Dimension pageSize = xmlSlideShow.getPageSize();
			is.close();
			String uidString = UUID.randomUUID().toString();//�ļ�����
			for (int i = 0; i < xslfSlides.length; i++) {
				System.out.print("��" + (i+1) + "ҳ��");
				setFont(xslfSlides[i]);
				BufferedImage img = new BufferedImage(pageSize.width*20,

				pageSize.height*20, BufferedImage.TYPE_INT_RGB);
			

				Graphics2D graphics = img.createGraphics();

				graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING,

						RenderingHints.VALUE_ANTIALIAS_ON);

				graphics.setRenderingHint(RenderingHints.KEY_RENDERING,

						RenderingHints.VALUE_RENDER_QUALITY);

				graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION,

						RenderingHints.VALUE_INTERPOLATION_BICUBIC);

				graphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS,

						RenderingHints.VALUE_FRACTIONALMETRICS_ON);

				graphics.setPaint(Color.white);
				graphics.scale(20, 20);// ��ͼƬ�Ŵ�times��
				graphics.fill(new Rectangle2D.Float(0, 0, pageSize.width*20,
				
				pageSize.height*20));
				

				xslfSlides[i].draw(graphics);
				
				FileOutputStream out = new FileOutputStream(path+"/"+ uidString + "."+picType);
				javax.imageio.ImageIO.write(img, "jpg", out);
				System.out.println(img);
				out.close();
				//���������ݿ�  roadshowid, uidString
				
				break;//�����ҪPPT�е�����ҳ�涼����ͼƬ����ôע�͵��þ���
			}
			System.out.println("...............PRINT_IMAGE_END.....");

			//�˴�Ӧ�ý� uidString ���ظ�ǰ̨��ʾ
			return uidString+".jpg";

		} catch (Exception e) {

		   e.printStackTrace();

		}


		return "fail";

	}
	
	//check the format of the file
	private static boolean checkFile(File file) {

		int pos = file.getName().lastIndexOf(".");

		String extName = "";

		if(pos >= 0) {

			 extName = file.getName().substring(pos);

		}

		if(".ppt".equalsIgnoreCase(extName) || ".pptx".equalsIgnoreCase(extName)){

			return true;

		}

		return false;

	} 


	
	

}
