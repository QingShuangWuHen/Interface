package com.zzy.Interface;

import org.testng.annotations.Test;

import com.alibaba.fastjson.JSONObject;

import org.testng.annotations.DataProvider;
import org.testng.annotations.BeforeClass;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Properties;
import java.util.Random;

import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.AfterClass;

public class Test_Interface {

	String filepath = System.getProperty("user.dir")+"\\lib\\InterfaceModel.xlsx";
	ResultMode rm = new ResultMode();

	@DataProvider(name = "dp")
	public Object[][] getObjects() {
		Object[][] testobj = null;
		try {
			testobj = ExcelUtils.getTableArray(filepath);
		} catch (Exception e) {
			// TODO 自动生成的 catch 块
			e.printStackTrace();
		}
		return testobj;
	}

	@Test(dataProvider = "dp")
	public void test_method(String line, String enable, String number, String name, String http, String path,
			String type, String input, String output, String error_code, String error_message,
			String result) {
			if (enable.equals("NO")) {
				System.out.print("测试用例不允许执行：" + number + name);
			} else {
				System.out.println("开始执行测试用例：" + number + name);
				if (type.equals("Post")) {
					try {
						output = postMethod(http + path, input);
					} catch (Exception e) {
						// TODO 自动生成的 catch 块
						e.printStackTrace();
					}

				} else if (type.equals("Get")) {
					try {
						output = getMethod(http + path);
					} catch (Exception e) {
						// TODO 自动生成的 catch 块
						e.printStackTrace();
					}

				}

				error_code = rm.getError_code();
				error_message = rm.getError_msg();
				if (error_code.equals("200")) {
					result = "Success";
				}else{
					result = "Fail";
				}
				
				System.out.println("output :" + output);
				System.out.println("error_code :" + error_code);
				System.out.println("error_message :" + error_message);
				System.out.println("result :" + result);
				
				try {
					ExcelUtils.writeExcel(filepath, line,
							output, error_code, error_message, result);
				} catch (Exception e) {
					// TODO 自动生成的 catch 块
					e.printStackTrace();
				}
			}

		
	}

	public String postMethod(String url, String jsonString) throws Exception {
		CloseableHttpClient httpClient = HttpClients.createDefault();
		HttpPost post = new HttpPost(url);

		StringEntity stringEntity = new StringEntity(jsonString);
		stringEntity.setContentType("text/json");
		post.setHeader("Content-Tpye", "application/json;charset=UTF-8");
		post.setEntity(stringEntity);

		CloseableHttpResponse response = httpClient.execute(post);
		String result = EntityUtils.toString(response.getEntity(), "UTF-8");
		// 以下内容根据需要修改，进行判断。
		JSONObject jsonObject = new JSONObject();
		JSONObject jsonObject2 = (JSONObject) jsonObject.parse(result);
		// System.out.println(jsonObject2);
		// System.out.println(jsonObject2.get("error_msg"));
		rm.setError_code(jsonObject2.getString("error_code"));
		rm.setError_msg(jsonObject2.getString("error_msg"));
		
		httpClient.close();
		response.close();

		return JSONObject.toJSONString(jsonObject2);
	}

	public String getMethod(String url) throws Exception {
		CloseableHttpClient httpClient = HttpClients.createDefault();
		HttpGet get = new HttpGet(url);

		CloseableHttpResponse response = httpClient.execute(get);
		String result = EntityUtils.toString(response.getEntity(), "UTF-8");
		// 以下内容根据需要修改，进行判断。
		JSONObject jsonObject = JSONObject.parseObject(result);
		// JSONObject jsonObject2 = (JSONObject)
		// jsonObject.getJSONObject("data").getJSONArray("uptime").get(0);
		// System.out.println(jsonObject2);
		// System.out.println(jsonObject2.get("app_name"));

		rm.setError_code(jsonObject.getString("error_code"));
		rm.setError_msg(jsonObject.getString("error_msg"));
		
		httpClient.close();
		response.close();

		return JSONObject.toJSONString(jsonObject);
	}
	
}
