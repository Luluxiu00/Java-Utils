/**
 * 
 */
package com.deppon.util;

import java.io.BufferedReader;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Map;

import org.apache.log4j.Logger;
 
/**
 *  Http请求的工具类
 * Created by WZ on 2017-5-24.
 * empCode: 326944
 */
public class HttpUtils {
 
	private static final Logger logger = Logger.getLogger(HttpUtils.class);
	
    private static final int TIMEOUT_IN_MILLIONS = 4000;
 
    public interface CallBack {
        void onRequestComplete(String result);
    }

    /**
	 * 向指定 URL 发送POST方法的请求
	 * 
	 * @param url 			发送请求的 URL
	 * @param param 	 	请求参数，请求参数应该是 name1=value1&name2=value2 的形式。
	 * @return	返回 post 请求结果 字符串
	 * @throws Exception
	 */
	public static String doPost(String url, String param) {
		PrintWriter out = null;
		BufferedReader in = null;
		try {
			URL realUrl = new URL(url);
			// 打开和URL之间的连接
			HttpURLConnection conn = (HttpURLConnection) realUrl.openConnection();
			// 设置通用的请求属性
			conn.setRequestProperty("accept", "*/*");
			conn.setRequestProperty("connection", "Keep-Alive");
			conn.setRequestMethod("POST");
			conn.setRequestProperty("Content-Type", "application/x-www-form-urlencoded");
			conn.setRequestProperty("charset", "UTF-8");
			conn.setUseCaches(false);
			// 发送POST请求必须设置如下两行
			conn.setDoOutput(true);
			conn.setDoInput(true);
			conn.setReadTimeout(TIMEOUT_IN_MILLIONS);
			conn.setConnectTimeout(TIMEOUT_IN_MILLIONS);

			if (StringUtils.isNotEmpty(param)) {
				// 发送请求参数
				out = new PrintWriter(conn.getOutputStream());
				out.print(param);
				out.flush();
			}
			// 定义BufferedReader输入流来读取URL的响应
			in = new BufferedReader(	new InputStreamReader(conn.getInputStream()));
			
			//响应结果
			StringBuffer resSb = new StringBuffer();
			String line = null;
			while ((line = in.readLine()) != null) {
				resSb.append( line) ;
			}
			return resSb.toString();
		} catch (Exception e) {
			logger.error("[ doPost ]，Params：<url, "+url+"><param, "+param+">。Error：" + e.getMessage());
		}finally {
			// 关闭资源
			try {
				if (out != null) 	out.close();
				if (in != null) 		in.close();
			} catch (IOException ex) {
				ex.printStackTrace();
			}
		}
		return null;
	}
	/**
	 * @param url				post请求url
	 * @param params		参数
	 * @return	返回 post 请求结果 字符串
	 */
	public static String doPost(String url, Map<String, String> params){
		//发送post请求
		return doPost(url, StringUtils.toString(params, "=", "&"));
	}
	
    /**
	 * @param url			get请求url
	 * @return				返回get请求结果
	 * @throws Exception
	 */
	public static String doGet(String url) {
		URL realUrl = null;
		HttpURLConnection conn = null;
		InputStream is = null;
		ByteArrayOutputStream baos = null;
		try {
			realUrl = new URL(url);
			conn = (HttpURLConnection) realUrl.openConnection();
			conn.setReadTimeout(TIMEOUT_IN_MILLIONS);
			conn.setConnectTimeout(TIMEOUT_IN_MILLIONS);
			conn.setRequestMethod("GET");
			conn.setRequestProperty("accept", "*/*");
			conn.setRequestProperty("connection", "Keep-Alive");
			
			int responseCode = conn.getResponseCode();
			//返回成功
			if ( responseCode == HttpURLConnection.HTTP_OK) {
				is = conn.getInputStream();
				baos = new ByteArrayOutputStream();

				int len = -1;
				byte[] buf = new byte[128];
				while ((len = is.read(buf)) != -1) {
					baos.write(buf, 0, len);
				}
				baos.flush();
				return baos.toString();
			} else {
				throw new RuntimeException(" Response Code is " + responseCode);
			}

		} catch (Exception e) {
			logger.error("[ doGet ]，Params：<url, "+url+"> Error：" + e.getMessage());
		} finally {
			try {
				if (is != null) 	is.close();
			} catch (IOException e) {}
			try {
				if (baos != null) 	baos.close();
			} catch (IOException e) {}
			conn.disconnect();
		}
		return null;
	}

	/**
	 * @param url		请求路径
	 * @return		返回get请求结果
	 */
	public static String doGet(String url, Map<String, String> params) {
		return doGet(url+"?"+StringUtils.toString(params, "=", "&"));
	}
	
	/**
	 * 异步的Get请求
	 * 
	 * @param urlStr
	 * @param callBack
	 */
	public static void doGetAsyn(final String urlStr, final CallBack callBack) {
		new Thread() {
			public void run() {
				try {
					String result = doGet(urlStr);
					if (callBack != null) {
						callBack.onRequestComplete(result);
					}
				} catch (Exception e) {
					e.printStackTrace();
				}
			};
		}.start();
	}
 
    /**
	 * 异步的Post请求
	 * 
	 * @param urlStr
	 * @param params
	 * @param callBack
	 * @throws Exception
	 */
	public static void doPostAsyn(final String urlStr, final String params, final CallBack callBack) throws Exception {
		new Thread() {
			public void run() {
				try {
					String result = doPost(urlStr, params);
					if (callBack != null) {
						callBack.onRequestComplete(result);
					}
				} catch (Exception e) {
					e.printStackTrace();
				}
			};
		}.start();
	}

}
