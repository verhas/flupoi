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
	private RangeProcessor lastAddedRangeProcessor;

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
		lastAddedRangeProcessor = new RangeProcessor(range, sheetName, processState);
		processorStack.add(lastAddedRangeProcessor);
		return this;
	}

	@Override
	public Workbook up() {
		lastAddedRangeProcessor.setDirection(RangeDirection.UP);
		return this;
	}

	@Override
	public Workbook down() {
		lastAddedRangeProcessor.setDirection(RangeDirection.DOWN);
		return this;
	}

	@Override
	public Workbook left() {
		lastAddedRangeProcessor.setDirection(RangeDirection.LEFT);
		return this;
	}

	@Override
	public Workbook right() {
		lastAddedRangeProcessor.setDirection(RangeDirection.RIGHT);
		return this;
	}

	@Override
	public Workbook until(Condition condition) {
		lastAddedRangeProcessor.setCondition(condition);
		return this;
	}

	@Override
	public Workbook until(String match) {
		lastAddedRangeProcessor.setCondition(new StringMatcherCondition(match));
		return this;
	}

	@Override
	public Workbook untilEmptyLine() {
		lastAddedRangeProcessor.setCondition(new EmptyLineMatcherCondition());
		return this;
	}

	@Override
	public Workbook untilEmptyCell() {
		lastAddedRangeProcessor.setCondition(new EmptyCellMatcherCondition());
		return this;
	}

	@Override
	public Workbook names(String[] names) {
		lastAddedRangeProcessor.setNames(names);
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

	private Map<Object, RangeProcessor> processorMap = new HashMap<>();

	@Override
	public Workbook target(Object key, Class<?> type) {
		processorMap.put(key, lastAddedRangeProcessor);
		lastAddedRangeProcessor.setTargetType(type);
		return this;
	}

	@Override
	public Workbook skip(int x, int y) {
		processorStack.add(new SkipProcessor(x, y, processState));
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
		RangeProcessor p = processorMap.get(key);
		return p == null ? null : p.getTargetCollection();
	}

}
