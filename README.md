#Table To Xls
## Preview
![HTML Table](table.png)

Result

![XLS Result](data.png)

## Usage
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