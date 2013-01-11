package javax0.flupoi;

import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

interface Processor {
	public void execute() throws InvalidFormatException, IOException,
	InstantiationException, IllegalAccessException,
	NoSuchFieldException, SecurityException;
}
