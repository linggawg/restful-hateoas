package wg.payroll.repository;

import org.springframework.data.jpa.repository.JpaRepository;
import wg.payroll.model.Employee;

public interface EmployeeRepository extends JpaRepository<Employee, Long> {

}