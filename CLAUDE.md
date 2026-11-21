# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

Talend component suite tFileExcel* for reading and writing Excel workbooks (Apache POI).
User-facing behaviour and parameters are described in `README.md` and the manuals in `doc/`.

## Build

```bash
mvn clean package                      # test, assemble component(s), copy them to ${custom_component_folder}
mvn test -Dtest=SomeTest#someMethod   # single test class or method
```

Java 11; `cimt-talendcomp-maven-plugin` 4.0. `custom_component_folder` is defined in `~/.m2/settings.xml`.
These tests reference local paths, servers or credentials on the author's machine and will likely fail elsewhere: `TestExcelComp`, `TestExcelSheetOutput`. Use `-DskipTests` when they are unavailable.

## Structure

- `src/main/java` (`de.jlo.talendcomp.excel`): the runtime library. Generated job code calls these classes by fully qualified name.
- `src/main/components/<Name>/`: `<Name>_java.xml` (descriptor), `*.javajet` (code-generation templates), `<Name>_messages.properties` (labels, `LONG_NAME` required) and the icon. A parameter change usually touches all three plus the library.
- Components: tFileExcelWorkbookOpen, tFileExcelWorkbookSave, tFileExcelSheetInput, tFileExcelSheetInputUnpivot, tFileExcelSheetOutput, tFileExcelSheetList, tFileExcelReferencedCellInput, tFileExcelReferencedCellOutput, tFileExcelNamedCellInput, tFileExcelNamedCellOutput. Each has its own plugin `<execution>` in `pom.xml`.
  Only tFileExcelWorkbookOpen bundles the jars; the others set `noJars` and need a jar-bundling component in the same job.
- Generated at `package`, do not hand-edit: the `<IMPORTS>` block and HEADER `VERSION`/`RELEASE_DATE` in `_java.xml`, and jars in the component folders. Bump the version in `pom.xml`.
