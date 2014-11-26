#Table To Xls
## Preview
![HTML Table](http://git.oschina.net/chyxion/table-to-xls/raw/master/html.png)

Result

![XLS Result](http://git.oschina.net/chyxion/table-to-xls/raw/master/xls.png)

## Usage
### Add Maven Repository
```xml
    <repository>
    	<id>chyxion-github</id>
    	<name>Chyxion Github</name>
    	<url>http://chyxion.github.io/maven/</url>
    </repository>
```
### Add Maven Dependency
```xml
    <dependency>
        <groupId>me.chyxion</groupId>
        <artifactId>table-to-xls</artifactId>
        <version>0.0.1-RELEASE</version>
    </dependency>
```
### Use In Code
```java
    StringBuilder html = new StringBuilder();
    Scanner s = new Scanner(
    	getClass().getResourceAsStream("/sample.html"), "utf-8");
    while (s.hasNext()) {
    	html.append(s.nextLine());
    }
    s.close();
    FileOutputStream fout = new FileOutputStream("data.xls");
    fout.write(TableToXls.process(html));
    fout.close();
```

## Contacts

chyxion@163.com
