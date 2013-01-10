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
private Processor lastAddedProcessor;
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
		lastAddedProcessor = new Processor(range, sheetName, processState);
		processorStack.add(lastAddedProcessor);
		return this;
	}

	@Override
	public Workbook up() {
		lastAddedProcessor.setDirection(RangeDirection.UP);
		return this;
	}

	@Override
	public Workbook down() {
		lastAddedProcessor.setDirection(RangeDirection.DOWN);
		return this;
	}

	@Override
	public Workbook left() {
		lastAddedProcessor.setDirection(RangeDirection.LEFT);
		return this;
	}

	@Override
	public Workbook right() {
		lastAddedProcessor.setDirection(RangeDirection.RIGHT);
		return this;
	}

	@Override
	public Workbook until(Condition condition) {
		lastAddedProcessor.setCondition(condition);
		return this;
	}

	@Override
	public Workbook until(String match) {
		lastAddedProcessor.setCondition(new StringMatcherCondition(match));
		return this;
	}

	@Override
	public Workbook untilEmptyLine() {
		lastAddedProcessor.setCondition(new EmptyLineMatcherCondition());
		return this;
	}

	@Override
	public Workbook untilEmptyCell() {
		throw new NotImplementedException("");
	}

	@Override
	public Workbook names(String[] names) {
		lastAddedProcessor.setNames(names);
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
		processorMap.put(key, lastAddedProcessor);
		lastAddedProcessor.setTargetType(type);
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
		processState.resetPositionPoint();
		processState.setIs(is);

		for (Processor processor : processorStack) {
			processor.execute();
		}
		return this;
	}

	@Override
	public Collection<?> fetch(Object key) {
		Processor p = processorMap.get(key);
		return p == null ? null : p.getTargetCollection();
	}

}
