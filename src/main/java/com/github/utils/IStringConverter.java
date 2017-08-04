package com.github.utils;

/**
 * 自定义字符串转换接口，用于将Excel导入的数据转换为自定义类型，例如将用';'区分的字符串转换为数组
 * @author XiaoYu
 *
 */
public interface IStringConverter {

	Object convert(String field, String value);
}
