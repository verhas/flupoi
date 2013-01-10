package javax0.flupoi;

class Coordinate {
	int value;
	boolean defined = false;
	private int relativity = 0;

	public int getValue() {
		return value;
	}

	public void setValue(int value) {
		this.value = value;
		this.defined = true;
	}

	public boolean isDefined() {
		return defined;
	}

	public void setDefined(boolean defined) {
		this.defined = defined;
	}

	public int getRelativity() {
		return relativity;
	}

	public void resetRelativity() {
		relativity = 0;
	}

	public void setRelativity(char sig) {
		relativity = sig == '+' ? 1 : (sig == '-' ? -1 : 0);
	}

	public boolean isRelative() {
		return relativity != 0;
	}

	public String toString() {
		return isDefined() ? ((isRelative() ? (getRelativity() > 0 ? "+" : "-")
				: "") + value) : "*";
	}
}