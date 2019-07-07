package com.cn;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.UUID;

import javax.imageio.ImageIO;

import org.apache.poi.hslf.HSLFSlideShow;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class testPPT2 {
	
	public static void main(String[] args) {

		List<String> list =  converPPTtoImage("C:\\Users\\ThinkPad\\Desktop\\Test.ppt", "C:\\Users\\ThinkPad\\Desktop\\img", "jpg", 20);
		System.out.println(list.get(0));
	}
	
	public static List<String> converPPTtoImage(String orignalPPTFileName, String targetImageFileDir,
	        String imageFormatNameString, int times) {
	    List<String> imgList = new ArrayList<>();
	    List<String> imgNamesList = new ArrayList<String>();// PPT转成图片后所有名称集合
	    FileInputStream orignalPPTFileInputStream = null;
	    FileOutputStream orignalPPTFileOutStream = null;
	    XMLSlideShow oneHSLFSlideShow = null;
	    //创建文件夹
	    createDirIfNotExist(targetImageFileDir);
	    try {
	        try {
	            orignalPPTFileInputStream = new FileInputStream(orignalPPTFileName);
	        } catch (FileNotFoundException e) {
	            e.printStackTrace();
	            return Collections.emptyList();
	        }
	        try {
	            oneHSLFSlideShow = new XMLSlideShow(orignalPPTFileInputStream);
	        } catch (IOException e) {
	            e.printStackTrace();
	            return Collections.emptyList();
	        }
	        // 获取PPT每页的大小（宽和高度）
	        Dimension onePPTPageSize = oneHSLFSlideShow.getPageSize();
	        // 获得PPT文件中的所有的PPT页面（获得每一张幻灯片）,并转为一张张的播放片
	        XSLFSlide[] pptPageSlideList = oneHSLFSlideShow.getSlides();
	        // 下面循环的主要功能是实现对PPT文件中的每一张幻灯片进行转换和操作
	        for (int i = 0; i < pptPageSlideList.length; i++) {
				setFont(pptPageSlideList[i]);
	            // 创建BufferedImage对象，图像的尺寸为原来的每页的尺寸*倍数times
	            BufferedImage oneBufferedImage = new BufferedImage(onePPTPageSize.width * times,
	                    onePPTPageSize.height * times, BufferedImage.TYPE_INT_RGB);
	            Graphics2D oneGraphics2D = oneBufferedImage.createGraphics();
	            // 设置转换后的图片背景色为白色
	            oneGraphics2D.setPaint(Color.white);
	            oneGraphics2D.scale(times, times);// 将图片放大times倍
	            oneGraphics2D
	                    .fill(new Rectangle2D.Float(0, 0, onePPTPageSize.width * times, onePPTPageSize.height * times));
	            pptPageSlideList[i].draw(oneGraphics2D);
	            // 设置图片的存放路径和图片格式，注意生成的图片路径为绝对路径，最终获得各个图像文件所对应的输出流对象
	            try {
	                String imgName = (i + 1) + "_" + UUID.randomUUID().toString() + "." + imageFormatNameString;
	                imgNamesList.add(imgName);// 将图片名称添加的集合中
	                imgList.add(imgName);
	                orignalPPTFileOutStream = new FileOutputStream(targetImageFileDir + imgName);

	            } catch (FileNotFoundException e) {
	                e.printStackTrace();
	                return Collections.emptyList();
	            }
	            // 转换后的图片文件保存的指定的目录中
	            try {
	                ImageIO.write(oneBufferedImage, imageFormatNameString, orignalPPTFileOutStream);
	            } catch (IOException e) {
	                e.printStackTrace();
	                return Collections.emptyList();
	            }
	           break;
	        }

	    } finally {
	        try {
	            if (orignalPPTFileInputStream != null) {
	                orignalPPTFileInputStream.close();
	            }
	        } catch (IOException e) {
	            e.printStackTrace();
	        }

	        try {
	            if (orignalPPTFileOutStream != null) {
	                orignalPPTFileOutStream.close();
	            }
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	    }

	    return imgList;
	}
	
	public static File createDirIfNotExist(String file) {
	    File fileDir = new File(file);
	    if (!fileDir.exists()) {
	        fileDir.mkdirs();
	    }
	    return fileDir;
	}
	
	
	private static void setFont(XSLFSlide slide) {
        for (XSLFShape shape : slide.getShapes()){
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
}