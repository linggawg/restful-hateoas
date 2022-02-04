package wg.payroll.repository;

import org.springframework.data.jpa.repository.JpaRepository;
import wg.payroll.model.Order;

public interface OrderRepository extends JpaRepository<Order, Long> {
}
