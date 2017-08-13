package com.cloudera.mapper;

import org.apache.hadoop.io.LongWritable;
import org.apache.hadoop.mapreduce.Mapper;
import org.apache.hadoop.io.Text;

import java.io.IOException;

public class ExcelStoreMapper extends Mapper<LongWritable, Text, Text, Text> {

	public void map(LongWritable key, Text value, Context context)
			throws InterruptedException, IOException {
		String line = value.toString();
		
		context.write(new Text(line), null);
	
	}
	
}
