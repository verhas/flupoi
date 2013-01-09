package javax0.flupoi;

import java.util.Collection;

import org.apache.poi.ss.usermodel.Cell;

public interface Condition {
	boolean match(Collection<Cell> cells);
}
