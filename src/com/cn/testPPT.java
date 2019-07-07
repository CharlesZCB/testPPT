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
		//实际就是传进来一个 文件 需要这个文件的Orignalname 属性值
		String infoString = doPPTtoImage2007(new File("C:\\Users\\ThinkPad\\Desktop\\Test.ppt"), "C:\\Users\\ThinkPad\\Desktop\\img", "a", "jpg");
		System.err.println(infoString);
	}

	/**
	 * 设置PPTX字体 
	 * @param slide
	 */
    private static void setFont(XSLFSlide slide) {
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTextShape) {
            	XSLFTextShape txtshape = (XSLFTextShape)shape ;
                for (XSLFTextParagraph paragraph : txtshape.getTextParagraphs()) {
                    List<XSLFTextRun> truns = paragraph.getTextRuns();
                    for (XSLFTextRun trun : truns) {
                    	trun.setFontFamily("宋体");
                    }
                }
            }
        }
    }
	

	public static String doPPTtoImage2007(File file,String path,String picName,String picType) {

		try {

			boolean isppt = checkFile(file);

			if (!isppt) {
				System.out.println("文件格式只能是PPT 或者 PPTX");
			return "format error";
			}

			FileInputStream is = new FileInputStream(file);

			XMLSlideShow xmlSlideShow = new XMLSlideShow(is);

			XSLFSlide[] xslfSlides = xmlSlideShow.getSlides();

			Dimension pageSize = xmlSlideShow.getPageSize();
			is.close();
			String uidString = UUID.randomUUID().toString();//文件名字
			for (int i = 0; i < xslfSlides.length; i++) {
				System.out.print("第" + (i+1) + "页。");
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
				graphics.scale(20, 20);// 将图片放大times倍
				graphics.fill(new Rectangle2D.Float(0, 0, pageSize.width*20,
				
				pageSize.height*20));
				

				xslfSlides[i].draw(graphics);
				
				FileOutputStream out = new FileOutputStream(path+"/"+ uidString + "."+picType);
				javax.imageio.ImageIO.write(img, "jpg", out);
				System.out.println(img);
				out.close();
				//保存在数据库  roadshowid, uidString
				
				break;//如果需要PPT中的所有页面都生成图片，那么注释掉该句子
			}
			System.out.println("...............PRINT_IMAGE_END.....");

			//此处应该将 uidString 返回给前台显示
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
