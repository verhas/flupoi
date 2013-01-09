package javax0.flupoi;

import java.io.IOException;
import java.io.InputStream;
import java.util.Collection;
import java.util.Deque;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.Map;
import java.util.StringTokenizer;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.eval.NotImplementedException;

public class WorkbookPOI implements Workbook {

	private ProcessState processState = new ProcessState();
	private InputStream is;
	private Deque<Processor> processorStack = new LinkedList<>();

	@Override
	public Workbook stream(InputStream is) {
		this.is = is;
		return this;
	}

	private String sheetName;

	@Override
	public Workbook sheet(String name) {
		sheetName = name;
		return this;
	}

	@Override
	public Workbook range(String range) {
		processorStack.push(new Processor(range, sheetName, processState));
		return this;
	}

	@Override
	public Workbook up() {
		processorStack.peek().setDirection(RangeDirection.UP);
		return this;
	}

	@Override
	public Workbook down() {
		processorStack.peek().setDirection(RangeDirection.DOWN);
		return this;
	}

	@Override
	public Workbook left() {
		processorStack.peek().setDirection(RangeDirection.LEFT);
		return this;
	}

	@Override
	public Workbook right() {
		processorStack.peek().setDirection(RangeDirection.RIGHT);
		return this;
	}

	@Override
	public Workbook until(Condition condition) {
		processorStack.peek().setCondition(condition);
		return this;
	}

	@Override
	public Workbook until(String match) {
		processorStack.peek().setCondition(new StringMatcherCondition(match));
		return this;
	}

	@Override
	public Workbook untilEmptyLine() {
		processorStack.peek().setCondition(new EmptyLineMatcherCondition());
		return this;
	}

	@Override
	public Workbook untilEmptyCell() {
		throw new NotImplementedException("");
	}

	@Override
	public Workbook names(String[] names) {
		processorStack.peek().setNames(names);
		return this;
	}

	@Override
	public Workbook names(String names) {
		StringTokenizer st = new StringTokenizer(names, ",");
		String[] namesArray = new String[st.countTokens()];
		for (int i = 0; st.hasMoreElements(); i++) {
			namesArray[i] = st.nextToken();
		}
		names(namesArray);
		return this;
	}

	private Map<Object, Processor> processorMap = new HashMap<>();

	@Override
	public Workbook target(Object key, Class<?> type) {
		processorMap.put(key, processorStack.peek());
		processorStack.peek().setTargetType(type);
		return this;
	}

	@Override
	public Workbook skip() {
		return this;
	}

	@Override
	public Workbook cell(String name, String range) {
		return range(range).names(name);
	}

	@Override
	public Workbook execute() throws InvalidFormatException, IOException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException {
		processState.setIs(is);
		for (Processor processor : processorStack) {
			processor.execute();
		}
		return this;
	}

	@Override
	public Collection<?> fetch(Object key) {
		return processorMap.get(key).getTargetCollection();
	}

}
