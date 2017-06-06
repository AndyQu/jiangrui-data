//EtlScript.extractAndAssemble(
//        ['/Users/andy/src/jiangrui-data/src/test/resources/test.xls'],
//        EtlScript.dataSpecification,
//        "output.xls"
//)

//EtlScript.extracFromFolder("/Users/andy/src/jiangrui-data/src/test/resources/12_21")
def File topFolder = new File("/Users/andy/Downloads/乙醇水冷凝0606/")
topFolder.listFiles().each {
    it->
        if(it.isDirectory()){
            println it
            EtlScript.extracFromFolder(it.absolutePath)
        }
}