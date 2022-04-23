#Table To Xls
## Preview
![HTML Table](/doc/html.png)

Result

![XLS Result](/doc/xls.png)

## Usage

### Add Maven Dependency
```xml
<dependency>
    <groupId>me.chyxion</groupId>
    <artifactId>table-to-xls</artifactId>
    <version>0.0.2</version>
</dependency>
```

### Use In Code
```java
TableToXls.process(getClass().getResourceAsStream("/sample.html"),
    StandardCharsets.UTF_8, "", new FileOutputStream("target/data.xlsx"));
```

## Contact

chyxion@163.com
