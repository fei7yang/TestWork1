package com.dfl.report.mfcadd;

import com.teamcenter.rac.kernel.TCComponentBOMLine;

public class NissanStation {
	private TCComponentBOMLine curLine;
	private TCComponentBOMLine[] Mfg0predecessors;
	private TCComponentBOMLine[] Mfg0successors;
	private int seqno = 0;
	private boolean leftRight = false;
	private String  name = "";
	private boolean passed = false;
	/**
	 * @return the curLine
	 */
	public TCComponentBOMLine getCurLine() {
		return curLine;
	}
	/**
	 * @param curLine the curLine to set
	 */
	public void setCurLine(TCComponentBOMLine curLine) {
		this.curLine = curLine;
	}
	/**
	 * @return the mfg0predecessors
	 */
	public TCComponentBOMLine[] getMfg0predecessors() {
		return Mfg0predecessors;
	}
	/**
	 * @param mfg0predecessors the mfg0predecessors to set
	 */
	public void setMfg0predecessors(TCComponentBOMLine[] mfg0predecessors) {
		Mfg0predecessors = mfg0predecessors;
	}
	/**
	 * @return the mfg0successors
	 */
	public TCComponentBOMLine[] getMfg0successors() {
		return Mfg0successors;
	}
	/**
	 * @param mfg0successors the mfg0successors to set
	 */
	public void setMfg0successors(TCComponentBOMLine[] mfg0successors) {
		Mfg0successors = mfg0successors;
	}
	/**
	 * @return the seqno
	 */
	public int getSeqno() {
		return seqno;
	}
	/**
	 * @param seqno the seqno to set
	 */
	public void setSeqno(int seqno) {
		this.seqno = seqno;
	}
	public boolean isLeftRight() {
		return leftRight;
	}
	public void setLeftRight(boolean leftRight) {
		this.leftRight = leftRight;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public boolean isPassed() {
		return passed;
	}
	public void setPassed(boolean passed) {
		this.passed = passed;
	}
	
}
