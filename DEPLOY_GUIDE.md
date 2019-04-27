## .m2/settings.xml
```xml
 <profile>
	 <id>gpg</id>
	 <properties>
		 <gpg.executable>/usr/local/bin/gpg</gpg.executable>
		 <gpg.passphrase>******</gpg.passphrase>
	 </properties>
</profile>
```

## 	mvn cli
```shell
 # clean package javadoc:jar source:jar gpg:sign deploy:deploy -P release -X
```

## 	Maven plugins
### pgp
```xml
<plugin>
	<groupId>org.apache.maven.plugins</groupId>
	<artifactId>maven-gpg-plugin</artifactId>
	<version>1.5</version>
	<executions>
		<execution>
			<phase>verify</phase>
			<goals>
				<goal>sign</goal>
			</goals>
		</execution>
	</executions>
</plugin>
```
### distribution
```xml
    <distributionManagement>
        <repository>
            <id>braisdom</id>
            <name>snapshots repository</name>
            <url>https://oss.sonatype.org/service/local/staging/deploy/maven2</url>
        </repository>
        <snapshotRepository>
            <id>snapshots</id>
            <name>Nexus Snapshot Repository</name>
            <url>https://oss.sonatype.org/content/repositories/snapshots</url>
        </snapshotRepository>
    </distributionManagement>

```