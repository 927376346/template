package com.liangbintao.action.base;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.ServletOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.struts2.ServletActionContext;
import org.apache.struts2.convention.annotation.Action;
import org.apache.struts2.convention.annotation.Namespace;
import org.apache.struts2.convention.annotation.ParentPackage;
import org.apache.struts2.convention.annotation.Result;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Scope;
import org.springframework.stereotype.Controller;

import com.liangbintao.domain.base.Area;
import com.liangbintao.service.base.AreaService;
import com.opensymphony.xwork2.ActionContext;
import com.opensymphony.xwork2.ActionSupport;
import com.opensymphony.xwork2.ModelDriven;

@Controller
@Scope("prototype")
@Namespace("/")
@ParentPackage("json-default")
public class AreaAction extends ActionSupport implements ModelDriven<Area> {

	private static final long serialVersionUID = 1L;
	// 模型注入
	private Area area = new Area();

	@Override
	public Area getModel() {
		return area;
	}

	@Autowired
	private AreaService areaService;

	private String file;
	private String fileFileName;
	private String fileContentType;

	public void setFile(String file) {
		this.file = file;
	}

	public void setFileFileName(String fileFileName) {
		this.fileFileName = fileFileName;
	}

	public void setFileContentType(String fileContentType) {
		this.fileContentType = fileContentType;
	}

	@Action(value = "area_batchImport", results = { @Result(name = "success", type = "json") })
	public String importXls() {
		String msg = "";
		try {
			Workbook workbook = null;
			// 判断后缀名
			if (fileFileName.endsWith(".xls")) {
				workbook = new HSSFWorkbook(new FileInputStream(file));
			} else if (fileFileName.endsWith(".xlsx")) {
				workbook = new XSSFWorkbook(new FileInputStream(file));
			}

			List<Area> areas = new ArrayList<Area>();
			// 1.创建workbook对象
			// 2.获取Sheet页
			Sheet sheet = workbook.getSheetAt(0);
			// 3.获取row行
			for (int i = 0; i < sheet.getLastRowNum(); i++) {
				// 获取每一行
				Row row = sheet.getRow(i);
				// 获取每一个单元格cell并且获取值
				String id = row.getCell(0).getStringCellValue();
				String province = row.getCell(1).getStringCellValue();
				String city = row.getCell(2).getStringCellValue();
				String district = row.getCell(3).getStringCellValue();
				String postcode = row.getCell(4).getStringCellValue();

				Area area = new Area(id, province, city, district, postcode);
				areas.add(area);
			}
			// 保存数据
			areaService.saveBatch(areas);
			// 成功
			msg = "成功";
		} catch (Exception e) {
			e.printStackTrace();
			msg = "失败";
		}
		ActionContext.getContext().getValueStack().push(msg);
		return SUCCESS;
	}

	/**
	 * 将数据转换为excel表格方式导出
	 * 
	 * @return
	 */
	@Action(value = "area_export")
	public String exportXls() {
		// 1.查询所有的数据
		List<Area> areas = areaService.findAll();
		
		// 2.将数据写入Excel文件
		// 2.1 创建一个workbook对象
		Workbook workbook = new HSSFWorkbook();
		// 规定10条数据一个sheet
		int size = areas.size();
		int pageSize = 100;
		// 总的sheet页数
		int count = (size % pageSize) == 0 ? size / pageSize : (size / pageSize + 1);

		for (int i = 0; i < count; i++) {
			// 2.2 创建 sheet
			Sheet sheet = workbook.createSheet("区域数据" + i);
			// 2.3 创建row表头
			Row row = sheet.createRow(0);
			row.createCell(0).setCellValue("区域编号");
			row.createCell(1).setCellValue("省份");
			row.createCell(2).setCellValue("城市");
			row.createCell(3).setCellValue("区域");
			row.createCell(4).setCellValue("邮编");
			// 2.4 遍历集合,创建数据行
			for (int j = i * pageSize + 1; j < (i + 1) * pageSize; j++) {
				// 当数据的记录数小于或等于j的时候,我们就停止创建数据行
				if (j >= size) {  
					break;
				}
				// 创建数据行
				Row row2 = sheet.createRow(sheet.getLastRowNum() + 1);
				// 给每行单元格填充数据
				row2.createCell(0).setCellValue(areas.get(j).getId());
				row2.createCell(1).setCellValue(areas.get(j).getProvince());
				row2.createCell(2).setCellValue(areas.get(j).getCity());
				row2.createCell(3).setCellValue(areas.get(j).getDistrict());
				row2.createCell(4).setCellValue(areas.get(j).getPostcode());
			}
		}
		
		try {
			// 3.将Excel写回客户端
			String filename = "qusj.xls";
			// 获取输出流
			ServletOutputStream stream = ServletActionContext.getResponse().getOutputStream();
			// 一个流,两个头
			ServletActionContext.getResponse().setContentType(
					ServletActionContext.getServletContext().getMimeType(filename));
			ServletActionContext.getResponse().setHeader("content-disposition","attachment;filename=" + filename);

			workbook.write(stream);
		} catch (IOException e) {
			e.printStackTrace();
		}
		return NONE;
	}
}
