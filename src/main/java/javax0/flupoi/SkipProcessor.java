package javax0.flupoi;

import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class SkipProcessor implements Processor {
	private ProcessState processState = new ProcessState();
	private int x, y;

	public ProcessState getProcessState() {
		return processState;
	}

	public void setProcessState(ProcessState processState) {
		this.processState = processState;
	}

	SkipProcessor(int x, int y, ProcessState processState) {
		this.x = x;
		this.y = y;
		this.setProcessState(processState);
	}

	@Override
	public void execute() throws InvalidFormatException, IOException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException {
		Point p = getProcessState().getLastProcessedPosition();
		int nx = Math.max(0, p.getX().getValue() + x);
		int ny = Math.max(0,p.getY().getValue() + y);
		p.getX().setValue(nx);
		p.getY().setValue(ny);
	}

}
