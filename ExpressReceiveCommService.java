package com.deppon.dpap.module.expressclass2.server.service.impl;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.struts2.ServletActionContext;

import com.deppon.cims.module.attachment.server.util.ExcelUtil;
import com.deppon.dpap.common.shared.context.UserContext;
import com.deppon.dpap.common.shared.util.UUIDUtils;
import com.deppon.dpap.module.authorization.shared.domain.UserEntity;
import com.deppon.dpap.module.basicmodule.shared.define.CimsDef;
import com.deppon.dpap.module.expressConfig2.shared.exception.ExpressConfig2Exception;
import com.deppon.dpap.module.expressclass2.server.dao.IExpressReceiveCommDao;
import com.deppon.dpap.module.expressclass2.server.service.IExpressReceiveCommService;
import com.deppon.dpap.module.expressclass2.shared.domain.ExpressReceiveCommEntity;
import com.deppon.dpap.module.expressclass2.shared.exception.Express2Exception;
import com.deppon.dpap.module.teamConfig.server.service.IPickSingStandardService;
import com.deppon.dpap.module.teamConfig.shared.domain.FreightereStandardEntity;
import com.esafenet.dll.FileDlpUtil;

/**
 * 类名：ExpressReceiveCommService
 * 
 * <pre>
 * @author 326944
 * @date 2015-08-11
 * @desc 必须继承父类
 * 必须生成serialVersionUID
 * </pre>
 */
public class ExpressReceiveCommService implements IExpressReceiveCommService {

	// 注入ReceivecommDao
	private IExpressReceiveCommDao expressReceiveCommDao;
	// 注入IPickSingStandardService，用以查询数据字典
	private IPickSingStandardService pickSingStandardService;
	/**
	 * 记录日志
	 */
	private static Logger logger = Logger.getLogger(ExpressReceiveCommService.class);
	
	public void setExpressReceiveCommDao(IExpressReceiveCommDao expressReceiveCommDao) {
		this.expressReceiveCommDao = expressReceiveCommDao;
	}
	
	public void setPickSingStandardService(IPickSingStandardService pickSingStandardService) {
		this.pickSingStandardService = pickSingStandardService;
	}
	
	/**
	 * 初始化本月，本周等时间
	 * @param sendRewardEntity
	 * @return
	 * @author 223955
	 * @date 2015-08-08
	 */
	public ExpressReceiveCommEntity initQueryDate(ExpressReceiveCommEntity entity){
		ExpressReceiveCommEntity entity1 = new ExpressReceiveCommEntity();
		SimpleDateFormat sdf=new SimpleDateFormat("yyyy-MM-dd");
		SimpleDateFormat sdss=new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        if("week".equals(entity.getBillDateType())){//如果为本周时间
        	 Calendar ca = Calendar.getInstance();
             ca.setTime(new Date());
        	//本周第几天，周未为第一天
        	int week = ca.get(Calendar.DAY_OF_WEEK);
        	logger.info("==============本周第几天，注意，周未为第一天："+week);
        	ca.add(Calendar.DATE, -week+2);
        	logger.info("==========周一的日期是几日："+sdf.format(ca.getTime()));
        	entity1.setSignTimeBegin(sdf.format(ca.getTime())+" 00:00:00");
        	entity1.setSignTimeEnd(sdss.format(new Date()));
        }
		return entity1;
	}
	/**
	 * 根据条件查询派件提成数据
	 * 
	 * @return List<ReceivecommEntity>
	 * @author 326944
	 * @date 2015-08-11
	 * @param entity
	 */
	@Override
	public List<ExpressReceiveCommEntity> queryAllExpressReceiveComm(ExpressReceiveCommEntity entity) {
		if("week".equals(entity.getBillDateType())){
			//初始化时间
			ExpressReceiveCommEntity en =this.initQueryDate(entity);
			entity.setSignTimeBegin(en.getSignTimeBegin());
			entity.setSignTimeEnd(en.getSignTimeEnd());
		}else{
			//界面选择时间
        	String startTime = entity.getSignTimeBegin();
        	//界面选择时间
        	String endTime = entity.getSignTimeEnd();
        	entity.setSignTimeBegin(startTime+" 00:00:00");
        	entity.setSignTimeEnd(endTime +" 23:59:59");
        	logger.info("======界面选择=====开始日期:"+startTime+";结束日期："+endTime);
        	logger.info("==============开始日期:"+entity.getSignTimeBegin()+";结束日期："+entity.getSignTimeEnd());
        }
		return expressReceiveCommDao.queryAllExpressReceiveComm(entity);
	}

	

	/**
	 * 查询条件查询派件提成数据条数
	 * 
	 * @param entity
	 * @return
	 * @author 326944
	 * @date 2015-08-11
	 */
	@Override
	public Long countQueryAllExpressReceiveComm(ExpressReceiveCommEntity entity) {
		if("week".equals(entity.getBillDateType())){
			//初始化时间
			ExpressReceiveCommEntity en =this.initQueryDate(entity);
			entity.setSignTimeBegin(en.getSignTimeBegin());
			entity.setSignTimeEnd(en.getSignTimeEnd());
		}else{
			//界面选择时间
        	String startTime = entity.getSignTimeBegin().substring(0, 10);
        	//界面选择时间
        	String endTime = entity.getSignTimeEnd().substring(0, 10);
        	entity.setSignTimeBegin(startTime+" 00:00:00");
        	entity.setSignTimeEnd(endTime +" 23:59:59");
        	logger.info("======界面选择=====开始日期:"+startTime+";结束日期："+endTime);
        	logger.info("==============开始日期:"+entity.getSignTimeBegin()+";结束日期："+entity.getSignTimeEnd());
        }
		return expressReceiveCommDao.countQueryAllExpressReceiveComm(entity);
	}

	/**
	 * 根据ID查询派件提成数据
	 * 
	 * @param entity
	 * @return
	 * @author 326944
	 * @date 2015-08-11
	 */
	@Override
	public ExpressReceiveCommEntity queryExpressReceiveCommById(ExpressReceiveCommEntity entity) {
		return expressReceiveCommDao.queryExpressReceiveCommById(entity);
	}

	/**
	 * 批量修改修改派件提成数据
	 * 
	 * @param entity
	 * @author 326944
	 * @date 2015-08-11
	 */
	@Override
	public void updateExpressReceiveComm(List<ExpressReceiveCommEntity> entityList) {
		UserEntity use = (UserEntity) UserContext.getCurrentUser();
		String psncode = use.getEmpEntity().getEmpCode();
		for (ExpressReceiveCommEntity ent : entityList) {
			FreightereStandardEntity fEntity = pickSingStandardService
					.queryPsnDeptInfo(ent.getPsnCode());
			// 根据工号修改姓名
			ent.setPsnName(fEntity.getPsnName());
			// 根据工号修改岗位名称
			ent.setJobName(fEntity.getPostion());
			
			ent.setModifyUser(psncode);
			ent.setModifyDate(new Date());
		}
		// 调用DAO
		expressReceiveCommDao.updateExpressReceiveComm(entityList);
	}




	/**
	 * 数据导入
	 * @throws ParseException 
	 */
	@Override
	public void importExpressReceiveCommData(File[] files, String[] fileNames) throws ParseException {
		String endName = fileNames[0].substring(fileNames[0].lastIndexOf("."));
		// 限制上传附件格式
		if (!".xlsx".equals(endName)) {
			throw new Express2Exception("只能上传【.xlsx】格式的Excel!");
		}
		// 根据文件获取Excel
		Workbook wb = getWorkbook(files[0]);
		// 获取sheet数据
		Sheet sheet = wb.getSheetAt(0);
		List<ExpressReceiveCommEntity> entityList = readExcel(sheet);
		for (ExpressReceiveCommEntity entity : entityList) {
			entity.setId(UUIDUtils.getUUID().toString());
			expressReceiveCommDao.mergeExpressReceiveComm(entity);
		}
	}

	/**
	 * 根据文件读取并解密Excel的静态方法
	 * 
	 * @param File
	 *            execlFile
	 * @return Workbook
	 * @author 326944
	 * @date 2015-05-07
	 * */
	public static Workbook getWorkbook(File execlFile) {
		Workbook wb = null;
		InputStream inputStream = null;
		// 解密excel
		String newPath = FileDlpUtil.getDecryptFile(execlFile.getPath());
		File fn = new File(newPath);
		try {
			inputStream = new FileInputStream(fn);
			// 根据版本选择创建Workbook的方式
			wb = new XSSFWorkbook(inputStream);
			// 关闭流
			inputStream.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return wb;
	}

	/**
	 * 读取导入的excel
	 * 
	 * @param sheet
	 * @return
	 * @throws ParseException 
	 */
	@SuppressWarnings("deprecation")
	public List<ExpressReceiveCommEntity> readExcel(Sheet sheet) throws ParseException {
		if (sheet == null) {
			// 抛出异常
			throw new ExpressConfig2Exception("选择的sheet在Excel中不存在!");
		}
		List<ExpressReceiveCommEntity> entityList = new ArrayList<ExpressReceiveCommEntity>();
		UserEntity use = (UserEntity) UserContext.getCurrentUser();
		String psncode = use.getEmpEntity().getEmpCode();
		SimpleDateFormat sd = new SimpleDateFormat("yyyy-MM-dd");
		/** 获得总列数 */
		int coloumNum = 0;
		try {
			// 定义从第一列开始读
			coloumNum = sheet.getRow(0).getPhysicalNumberOfCells();
		} catch (Exception e) {
			throw new ExpressConfig2Exception("请检查导入EXCEL是否正确或是否有数据!");
		}
		// 定义从第二行开始读
		int rowNum = 1;
		while (rowNum >= 1) {
			Row row = sheet.getRow(rowNum);
			ExpressReceiveCommEntity entity = new ExpressReceiveCommEntity();
			// String code="";
			for (int cellNum = 0; cellNum < coloumNum; cellNum++) {
				// 从第一列开始
				if (row == null) {
					rowNum = -1;
					entity = null;
					break;
				}
				Cell cell = row.getCell(cellNum);
				String cellValue = "";
				if (cell != null) {
					cell.setCellType(Cell.CELL_TYPE_STRING);
					// 防止poi把单元格中的数字当做double来处理
					cellValue = cell.getStringCellValue().trim();
					// 事业部
					if (cellNum == 0) {
						if (!"".equals(cellValue) && cellValue != null) {
							entity.setCareerDept(cellValue);
						}
					}
					// 大区
					if (cellNum == CimsDef.one) {
						if (!"".equals(cellValue) && cellValue != null) {
							entity.setAreaDept(cellValue);
						}
					}
					// 计费日期
					if (cellNum == CimsDef.two) {
						if ("".equals(cellValue) || cellValue == null) {
							throw new ExpressConfig2Exception("excel中的计费日期字段为空!");
						} else {
							entity.setBillDate(sd.parse(cellValue));
						}
					}
					// 部门名称
					if (cellNum == CimsDef.three) {
						if (!"".equals(cellValue) && cellValue != null) {
							entity.setDeptName(cellValue);
						}
					}
					// 部门编码
					if (cellNum == CimsDef.four) {
						if (!"".equals(cellValue) && cellValue != null) {
							entity.setDeptCode(cellValue);
						}
					}
					// 运单号
					if (cellNum == CimsDef.five) {
						if ("".equals(cellValue) || cellValue == null) {
							throw new ExpressConfig2Exception("excel中的运单号字段为空!");
						} else {
							entity.setBillNo(cellValue);
						}
					}
					// 工号
					if (cellNum == CimsDef.six) {
						if ("".equals(cellValue) || cellValue == null) {
							throw new ExpressConfig2Exception("excel中的工号字段为空!");
						} else {
							entity.setPsnCode(cellValue);
						}
					}
					// 姓名
					if (cellNum == CimsDef.seven) {
						if ("".equals(cellValue) || cellValue == null) {
							throw new ExpressConfig2Exception("excel中的姓名字段为空!");
						} else {
							entity.setPsnName(cellValue);
						}
					}
					// 岗位名称
					if (cellNum == CimsDef.eight) {
						if ("".equals(cellValue) || cellValue == null) {
							throw new ExpressConfig2Exception("excel中的岗位名称字段为空!");
						} else {
							entity.setJobName(cellValue);
						}
					}
					// 业务类型
					if (cellNum == CimsDef.nine) {
						if ("".equals(cellValue) || cellValue == null) {
							throw new ExpressConfig2Exception("excel中的业务类型字段为空!");
						} else {
							entity.setTypeName(cellValue);
						}
					}
					// 产品类型
					if (cellNum == CimsDef.ten) {
						if ("".equals(cellValue) || cellValue == null) {
							throw new ExpressConfig2Exception("excel中的产品类型字段为空!");
						} else {
							entity.setTransportType(cellValue);
						}
					}
					// 开单金额
					if (cellNum == CimsDef.eleven) {
						if ("".equals(cellValue) || cellValue == null) {
							throw new ExpressConfig2Exception("excel中的开单金额字段为空!");
						} else {
							entity.setBillSum(cellValue);
						}
					}
					// 货物重量
					if (cellNum == CimsDef.twelve) {
						if ("".equals(cellValue) || cellValue == null) {
							throw new ExpressConfig2Exception("excel中的货物重量字段为空!");
						} else {
							entity.setGoodsWeight(cellValue);
						}
					}
					// 代收货款
					if (cellNum == CimsDef.thirteen) {
						if ("".equals(cellValue) || cellValue == null) {
							throw new ExpressConfig2Exception("excel中的代收货款字段为空!");
						} else {
							entity.setCollectionSum(cellValue);
						}
					}
					// 始发城市
					if (cellNum == CimsDef.fourteen) {
						if ("".equals(cellValue) || cellValue == null) {
							throw new ExpressConfig2Exception("excel中的始发城市字段为空!");
						} else {
							entity.setSetoutCity(cellValue);
						}
					}
					// 城市类别
					if (cellNum == CimsDef.fifteen) {
						if ("".equals(cellValue) || cellValue == null) {
							throw new ExpressConfig2Exception("excel中的城市类别字段为空!");
						} else {
							entity.setCityType(cellValue);
						}
					}
					// 到达城市
					if (cellNum == CimsDef.sixteen) {
						if ("".equals(cellValue) || cellValue == null) {
							throw new ExpressConfig2Exception("excel中的到达城市字段为空!");
						} else {
							entity.setReceiveCity(cellValue);
						}
					}
					// 收件标准
					if (cellNum == CimsDef.seventeen) {
						if ("".equals(cellValue) || cellValue == null) {
							throw new ExpressConfig2Exception("excel中的收件标准字段为空!");
						} else {
							entity.setReceiveStandard(cellValue);
							
						}
					}
					// 提成奖金
					if (cellNum == CimsDef.eighteen) {
						if ("".equals(cellValue) || cellValue == null) {
							throw new ExpressConfig2Exception("excel中的提成奖金字段为空!");
						} else {
							entity.setReward(cellValue);
						}
					}
					// 发货人客户编码
					if (cellNum == CimsDef.nineteen) {
						if ("".equals(cellValue) || cellValue == null) {
							throw new ExpressConfig2Exception("excel中的发货人客户编码字段为空!");
						} else {
							entity.setCustomerCoding(cellValue);
						}
					}
					// 修改原因
					if (cellNum == CimsDef.twenty) {
						if (!"".equals(cellValue) && cellValue != null) {
							entity.setModifyReason(cellValue);
						}
					}
				}else{
					throw new ExpressConfig2Exception("excel中第"+(rowNum+1)+"行,第"+(cellNum+1)+"列不能为空!"); 
				}
			}
			if (entity != null) {// 设置值
				entity.setModifyUser(psncode);
				entity.setModifyDate(new Date());
				entityList.add(entity);
			}
			rowNum++;
		}
		return entityList;
	}


	/**
	 * 批量导出
	 * 
	 * @param receivecommEntity
	 *            传入查询的entity
	 * @return 返回excel的名字
	 * 
	 */
	@Override
	public String excelExplore(ExpressReceiveCommEntity entity) {
		if("week".equals(entity.getBillDateType())){
			//初始化时间
			ExpressReceiveCommEntity en =this.initQueryDate(entity);
			entity.setSignTimeBegin(en.getSignTimeBegin());
			entity.setSignTimeEnd(en.getSignTimeEnd());
		}else{
			//界面选择时间
        	String startTime = entity.getSignTimeBegin();
        	//界面选择时间
        	String endTime = entity.getSignTimeEnd();
        	entity.setSignTimeBegin(startTime+" 00:00:00");
        	entity.setSignTimeEnd(endTime +" 23:59:59");
        	logger.info("======界面选择=====开始日期:"+startTime+";结束日期："+endTime);
        	logger.info("==============开始日期:"+entity.getSignTimeBegin()+";结束日期："+entity.getSignTimeEnd());
        }
		String fileName = "expressReceiveComm" + CimsDef.fileType;
		// 创建一个excel工作薄
		Workbook wb = new SXSSFWorkbook(CimsDef.rewardStandard);//1000
		// 导出excel的名称
		// 明细excel处理
		accruedExcel(wb, entity);
		// excel保存路径
		OutputStream os = null;
		try {
			HttpServletResponse response = ServletActionContext.getResponse();
			response.setContentType("application/x-msdownload");
			response.setHeader("Content-Disposition", "attachment; filename="
					+ fileName);
			os = response.getOutputStream();
			// 将excel写入到输出流
			wb.write(os);
		} catch (IOException e) {
			e.printStackTrace();
			throw new Express2Exception("导出文件生成失败!");
		} finally {
			try {
				if (os != null) {
					os.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
				throw new Express2Exception("导出文件输出流关闭失败!");
			}
		}
		return fileName;
	}

	/**
	 * 方法体说明：创建人员与业务量反馈信息的excel
	 * 
	 * @param wb
	 *            创建的excel对象
	 * @author 223955
	 * @param receivecommEntityList
	 */
	@SuppressWarnings("deprecation")
	private void accruedExcel(Workbook wb,ExpressReceiveCommEntity entity) {
		SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
		SimpleDateFormat formatss = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		Sheet hs = wb.createSheet("快递收件提成报表");// 创建一个sheet
		Row row = null;// 定义行
		CellStyle cellStyle = null;// 定义单元格样式
		CellStyle cellStyle1 = null;// 定义单元格样式
		Cell cell = null;// 定义单元格
		Font font = null;// 定义字体
		CellRangeAddress cellRangeAddress = null;// 定义合并单元格
		// 布局
		cellStyle = wb.createCellStyle();// 创建单元格样式
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);// 垂直
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);// 水平
		// 布局
		cellStyle1 = wb.createCellStyle();// 创建单元格样式
		font = wb.createFont();
		font.setFontHeightInPoints((short) CimsDef.twelve); // 字体大小
		cellStyle1.setFont(font);
		cellStyle1.setVerticalAlignment(CellStyle.VERTICAL_CENTER);// 垂直
		cellStyle1.setAlignment(CellStyle.ALIGN_CENTER);// 水平
		
		// 单元格合并 四个参数分别是：起始行，结束行，起始列，结束列
		// 表头合并
		cellRangeAddress = new CellRangeAddress(0, 0, 0, CimsDef.twentyTwo);//22
		hs.addMergedRegion(cellRangeAddress);
		// 第一行
		row = hs.createRow(0);// 在索引0的位置创建行（最顶端的行）
		cell = row.createCell(0); // 在索引0的位置创建单元格（左上端）
		cell.setCellType(Cell.CELL_TYPE_STRING); // 定义单元格为字符串类型
		cell.setCellStyle(cellStyle1);// 设置单元格样式
		cell.setCellValue("快递员收件提成报表");// 设置单元格值
		// 第二行
		row = hs.createRow(1);// 在索引1的位置创建行
		//row.setHeight((short) 500);// 设置行高度
		int width = CimsDef.excelWidth;
		ExcelUtil.createColumn(hs, row, cellStyle, 0, width, "序号");// 第一列序号
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.one, width, "事业部");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.two, width, "大区");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.three, width, "计费日期");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.four, width, "部门名称");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.five, width, "部门编码");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.six, width, "运单号");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.seven, width, "工号");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.eight, width, "姓名");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.nine, width, "岗位名称");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.ten, width, "业务类型");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.eleven, width, "产品类型");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.twelve, width, "开单金额");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.thirteen, width, "货物重量");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.fourteen, width, "代收货款");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.fifteen, width, "始发城市");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.sixteen, width, "城市类别");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.seventeen, width, "到达城市");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.eighteen, width, "收件标准");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.nineteen, width, "提成奖金");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.twenty, width, "发货人客户编码");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.twentyOne, width, "修改原因");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.twentyTwo, width, "修改人");
		ExcelUtil.createColumn(hs, row, cellStyle, CimsDef.twentyThree, width, "修改时间");
		// 数据库中存储的数据行
		int pagesize = CimsDef.limit;//10000
		// 查询导出数据的总数
		Long pageAll = expressReceiveCommDao.countQueryAllExpressReceiveComm(entity);
		// 根据行数求数据提取次数  
		int exporttimes = (int) (pageAll % pagesize > 0 ? pageAll / pagesize + 1 : pageAll / pagesize); 
		// 按次数将数据写入文件  
		for(int i = 0;i < exporttimes;i++){
			// 设置分页大小
			entity.setStart(i*pagesize);
			entity.setLimit(pagesize);
			// 调用Service
			List<ExpressReceiveCommEntity> entityList = expressReceiveCommDao.queryAllExpressReceiveComm(entity);
			int len = entityList.size() < pagesize ? entityList.size() : pagesize;
			for (int j = 0; j < len; j++) {
				ExpressReceiveCommEntity aq = entityList.get(j);
				row = hs.createRow(i * pagesize + j + 2);// 创建一行
				//给每一列赋值
				String billDate = "";//开单日期
				String modifydate = "";//修改日期

				if(aq.getBillDate()!=null){
					billDate = format.format(aq.getBillDate());//开单日期
				}
				if(aq.getModifyDate()!=null){
					modifydate = formatss.format(aq.getModifyDate());//修改日期
				}			
				ExcelUtil.setCellValue(0, row, cellStyle, String.valueOf(i * pagesize + j + 1));// 序号
				ExcelUtil.setCellValue(1, row, cellStyle, aq.getCareerDept());// 事业部
				ExcelUtil.setCellValue(CimsDef.two, row, cellStyle, aq.getAreaDept());// 大区
				ExcelUtil.setCellValue(CimsDef.three, row, cellStyle, billDate);//开单日期->计费日期
				ExcelUtil.setCellValue(CimsDef.four, row, cellStyle, aq.getDeptName());//部门名称
				ExcelUtil.setCellValue(CimsDef.five, row, cellStyle,aq.getDeptCode() );// 部门编码
				ExcelUtil.setCellValue(CimsDef.six, row, cellStyle, aq.getBillNo());// 运单号
				ExcelUtil.setCellValue(CimsDef.seven, row, cellStyle,  aq.getPsnCode());// 工号
				ExcelUtil.setCellValue(CimsDef.eight, row, cellStyle, aq.getPsnName());// 姓名
				ExcelUtil.setCellValue(CimsDef.nine, row, cellStyle, aq.getJobName());// 岗位名称
				ExcelUtil.setCellValue(CimsDef.ten, row, cellStyle, aq.getTypeName());// 业务类型
				ExcelUtil.setCellValue(CimsDef.eleven, row, cellStyle, aq.getTransportType());// 产品类型
				ExcelUtil.setCellValue(CimsDef.twelve, row, cellStyle, aq.getBillSum());// 开单金额
				ExcelUtil.setCellValue(CimsDef.thirteen, row, cellStyle, aq.getGoodsWeight());// 货物重量
				ExcelUtil.setCellValue(CimsDef.fourteen, row, cellStyle, aq.getCollectionSum());// 代收货款
				ExcelUtil.setCellValue(CimsDef.fifteen, row, cellStyle, aq.getSetoutCity());// 始发城市
				ExcelUtil.setCellValue(CimsDef.sixteen, row, cellStyle, aq.getCityType());// 城市类别
				ExcelUtil.setCellValue(CimsDef.seventeen, row, cellStyle, aq.getReceiveCity());// 到达城市
				ExcelUtil.setCellValue(CimsDef.eighteen, row, cellStyle, aq.getReceiveStandard());// 收件标准
				ExcelUtil.setCellValue(CimsDef.nineteen, row, cellStyle, aq.getReward());// 提成奖金
				ExcelUtil.setCellValue(CimsDef.twenty, row, cellStyle,  aq.getCustomerCoding());// 发货人客户编码
				ExcelUtil.setCellValue(CimsDef.twentyOne, row, cellStyle, aq.getModifyReason());// 修改原因
				ExcelUtil.setCellValue(CimsDef.twentyTwo, row, cellStyle, aq.getModifyUser());// 修改人
				ExcelUtil.setCellValue(CimsDef.twentyThree, row, cellStyle, modifydate);// 最后修改时间
			}
			entityList.clear();
		}
	}
	



}