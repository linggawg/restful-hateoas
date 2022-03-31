package wg.payroll.controller;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.hateoas.CollectionModel;
import org.springframework.hateoas.EntityModel;
import org.springframework.hateoas.IanaLinkRelations;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import wg.payroll.component.EmployeeModelAssembler;
import wg.payroll.exception.EmployeeNotFoundException;
import wg.payroll.model.Employee;
import wg.payroll.repository.EmployeeRepository;
import wg.payroll.service.ExportService;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

import static org.springframework.hateoas.server.mvc.WebMvcLinkBuilder.linkTo;
import static org.springframework.hateoas.server.mvc.WebMvcLinkBuilder.methodOn;

@RestController
public class EmployeeController {

    private final EmployeeRepository repository;

    private final EmployeeModelAssembler assembler;

    private final ExportService exportService;

    EmployeeController(EmployeeRepository repository, EmployeeModelAssembler assembler, ExportService exportService) {
        this.repository = repository;
        this.assembler = assembler;
        this.exportService = exportService;
    }

    @GetMapping("/employees")
    public CollectionModel<EntityModel<Employee>> all() {

        List<EntityModel<Employee>> employees = repository.findAll().stream()
                .map(assembler::toModel)
                .collect(Collectors.toList());
        return CollectionModel.of(employees, linkTo(methodOn(EmployeeController.class).all()).withSelfRel());
    }

    @PostMapping("/employees")
    public ResponseEntity<?> newEmployee(@RequestBody Employee newEmployee) {

        EntityModel<Employee> entityModel = assembler.toModel(repository.save(newEmployee));
        return ResponseEntity
                .created(entityModel.getRequiredLink(IanaLinkRelations.SELF).toUri())
                .body(entityModel);
    }

    @GetMapping("/employees/{id}")
    public EntityModel<Employee> one(@PathVariable Long id) {

        Employee employee = repository.findById(id)
                .orElseThrow(() -> new EmployeeNotFoundException(id));
        return assembler.toModel(employee);
    }

    @PutMapping("/employees/{id}")
    public ResponseEntity<?> replaceEmployee(@RequestBody Employee newEmployee, @PathVariable Long id) {

        Employee updatedEmployee = repository.findById(id)
                .map(employee -> {
                    employee.setName(newEmployee.getName());
                    employee.setRole(newEmployee.getRole());
                    return repository.save(employee);
                })
                .orElseGet(() -> {
                    newEmployee.setId(id);
                    return repository.save(newEmployee);
                });

        EntityModel<Employee> entityModel = assembler.toModel(updatedEmployee);
        return ResponseEntity
                .created(entityModel.getRequiredLink(IanaLinkRelations.SELF).toUri())
                .body(entityModel);
    }

    @DeleteMapping("/employees/{id}")
    public ResponseEntity<?> deleteEmployee(@PathVariable Long id) {

        repository.deleteById(id);
        return ResponseEntity.noContent().build();
    }

    @GetMapping("/employees/export")
    public ResponseEntity<HttpServletResponse> exportEmployee(HttpServletResponse response) {
        try{
            List<Employee> employees = repository.findAll();
            List<String> headers = Arrays.asList("First Name", "Last Name", "Role");

            response.setContentType("application/octet-stream");
            String headerKey = "Content-Disposition";
            String headerValue = "attachment; filename=Employee.xlsx";
            response.setHeader(headerKey, headerValue);

            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("Data");

            XSSFFont font = workbook.createFont();
            font.setFontHeightInPoints((short) 11);
            font.setFontName("Calibri");
            font.setColor(IndexedColors.BLACK.getIndex());

            CellStyle style = workbook.createCellStyle();
            style.setFont(font);

            XSSFRow rowHeader = sheet.createRow(0);
            for(String header : headers){
                exportService.createCell(rowHeader, headers.indexOf(header), header, style);
            }

            for(Employee employee : employees){
                int indexRow = employees.indexOf(employee);
                XSSFRow rowValue = sheet.createRow(indexRow+1);

                int indexCell = 0;
                exportService.createCell(rowValue, indexCell++, employee.getFirstName(), style);
                exportService.createCell(rowValue, indexCell++, employee.getLastName(), style);
                exportService.createCell(rowValue, indexCell, employee.getRole(), style);
            }

            for(String header : headers){
                sheet.autoSizeColumn(headers.indexOf(header));
            }

            ServletOutputStream outputStream = response.getOutputStream();
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
            return ResponseEntity.status(HttpStatus.OK).body(response);
        }catch (Exception e){
            e.printStackTrace();
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(response);
        }
    }
}
