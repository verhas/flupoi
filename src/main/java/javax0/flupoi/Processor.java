package javax0.flupoi;

import java.io.IOException;
import java.lang.reflect.Field;
import java.util.Collection;
import java.util.LinkedList;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;

public class Processor {
	private Range range;
	private String sheetName;

	private ProcessState processState;

	protected Processor(String range, String sheetName,
			ProcessState processState) {
		this.range = new Range(range);
		this.setSheetName(sheetName);
		this.setProcessState(processState);
	}

	public void setDirection(RangeDirection direction) {
		range.setDirection(direction);
	}

	private Condition condition;

	public Range getRange() {
		return range;
	}
	
	public Condition getCondition() {
		return condition;
	}

	/**
	 * Set the condition to be used by range iterators to stop.
	 * 
	 * @param condition
	 */
	public void setCondition(Condition condition) {
		this.condition = condition;
	}

	private String[] names;

	public void setNames(String[] names) {
		this.names = names;
	}

	private Class<?> targetType;

	public void setTargetType(Class<?> targetType) {
		this.targetType = targetType;
	}

	private Collection<Object> targetCollection;

	public Collection<?> getTargetCollection() {
		return targetCollection;
	}

	private void storeCell(Object target, String name, Object value)
			throws NoSuchFieldException, SecurityException,
			IllegalArgumentException, IllegalAccessException {
		if (target instanceof Map) {
			@SuppressWarnings("unchecked")
			Map<String, Object> map = (Map<String, Object>) target;
			map.put(name, value);
		} else {
			Field field = target.getClass().getField(name);
			field.setAccessible(true);
			field.set(target, value);
		}
	}

	private Object executeOneRow(Collection<Cell> cells)
			throws InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException {
		Object target = targetType.newInstance();
		int i = 0;
		for (Cell cell : cells) {
			String name = names[i++];
			storeCell(target, name, cell.getStringCellValue());
		}
		return target;
	}

	public void execute() throws InvalidFormatException, IOException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException {
		RangeIterator cellCollections = new RangeIterator(this);
		getProcessState().setRangeIterator(cellCollections);
		targetCollection = new LinkedList<>();
		for (Collection<Cell> cells : cellCollections) {
			targetCollection.add(executeOneRow(cells));
		}
	}

	public String getSheetName() {
		return sheetName;
	}

	private void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public ProcessState getProcessState() {
		return processState;
	}

	private void setProcessState(ProcessState processState) {
		this.processState = processState;
	}
}
