package com.grails.plugins.xlsarea

class XlsArea {
	def startsWith
	def content = []
	def endsWith
	def matchResult
	def xls
	def contents = [:]
	def key
	final String CLEAR_KEY = ""

	void setSource(source){
		xls = source
		def i = 0, j;

		def enterIndex = -1
		for(j = 0; j < source.size(); j++) {
			def enterStep = startsWith[i]
			def sourceLine = source[j]
			if(checkMatch(sourceLine, enterStep)){
				i ++;
				if(i == startsWith.size()){
					enterIndex = j;
					break;
				}
			} else {
				i = 0
			}
		}

		if(enterIndex != -1){
			i = 0; // index de saida
			for(j = enterIndex + 1; j < source.size(); j++) {
				def sourceLine = source[j]
				def exitStep = endsWith[i]
				
				if(checkMatch(sourceLine, exitStep)){
					i ++;
					if(i == endsWith.size()){
						break;
					}
				}
				content.push(sourceLine)
			}
		}
		if(key){
			contents.get(key, content)
			key = CLEAR_KEY
		}
		
	}

	def loadContent(){
		if(xls){
			content = []
			setSource(xls)
			return this
		}
	}

	def startsWith(params){
		startsWith = params
		return this
	}

	def endsWith(params){
		endsWith = params
		return this
	}

	void setContentLabel(contentKey){
		key = contentKey
	}


	/**
	*	Verifica se alguma palavra da linha casa, baseado na palavra passada. Caso a palavra seja 'blankline' ou uma lista vazia ([])
	*	o pattern será nulo para que dê match com linha vazia.
	*	Sendo o fonte uma linha vazia colocamos o fonte como a string "nula" para dar match com "blankline" ou [] se for o caso.
	*
	*/
	
	def checkMatch(sourceRow, word) {
		def pattern = (word == 'blankline' || word instanceof List) ? ~/null/ : ~word
		def find = false
		if(sourceRow && sourceRow[0] != "") {
			for(def i = 0 ; i < sourceRow.size() ; i++) {
				def source = String.valueOf(sourceRow[i])
				if(pattern.matcher(source).matches()) {
					find = true
					break
				}
			}
		} else {
			def source = (sourceRow instanceof List) ? "null" : "do nothing"
			find = pattern.matcher(source).matches()
		}

		return find
	}

}