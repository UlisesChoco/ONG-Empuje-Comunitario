package com.grupo_c.SistemasDistribuidosTP;

import com.grupo_c.SistemasDistribuidosTP.constants.RoleEnum;
import com.grupo_c.SistemasDistribuidosTP.entity.Donation;
import com.grupo_c.SistemasDistribuidosTP.entity.Inventory;
import com.grupo_c.SistemasDistribuidosTP.entity.Role;
import com.grupo_c.SistemasDistribuidosTP.entity.User;
import com.grupo_c.SistemasDistribuidosTP.repository.IDonationRepository;
import com.grupo_c.SistemasDistribuidosTP.service.IDonationService;
import com.grupo_c.SistemasDistribuidosTP.service.IRoleService;
import com.grupo_c.SistemasDistribuidosTP.service.IUserService;
import jakarta.transaction.Transactional;
import org.springframework.boot.CommandLineRunner;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.stereotype.Component;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

@Component
public class ApplicationRunner implements CommandLineRunner {
    private final IRoleService roleService;
    private final IUserService userService;
    private final PasswordEncoder passwordEncoder;
    private final IDonationRepository donationService;
    public ApplicationRunner(IRoleService roleService, IUserService userService, PasswordEncoder passwordEncoder, IDonationRepository donationService) {
        this.roleService = roleService;
        this.userService = userService;
        this.passwordEncoder = passwordEncoder;
        this.donationService = donationService;
    }
    @Override
    @Transactional
    public void run(String... args) throws Exception {
        loadRoles();
        loadUser();
        loadDonations();
    }
    private void loadRoles() {
        if(roleService.findAll().isEmpty()) {
            Set<Role> roles = new HashSet<>();
            roles.add(new Role(RoleEnum.VOLUNTARIO.name()));
            roles.add(new Role(RoleEnum.COORDINADOR.name()));
            roles.add(new Role(RoleEnum.VOCAL.name()));
            roles.add(new Role(RoleEnum.PRESIDENTE.name()));
            roleService.saveAll(roles);
        }
    }
    private void loadUser() {
        if(userService.findAll().isEmpty()) {
            Role role = roleService.findByName(RoleEnum.PRESIDENTE.name());
            User user = new User();
            user.setUsername("Presidente");
            user.setName("Tomás");
            user.setSurname("Lopez");
            user.setPhoneNumber("+54 11 4376-3389");
            user.setPassword(passwordEncoder.encode("admin"));
            user.setEmail("tomaslopez1987@gmail.com");
            user.setIsActive(true);
            user.setRoles(Set.of(role));
            userService.save(user);
        }
    }
    private void loadDonations() {
        if (donationService.findAll().isEmpty()) {
            List<Donation> donations = new ArrayList<>();
            LocalDate today = LocalDate.now();

            Donation remera = new Donation();
            remera.setCategory(Inventory.Category.ROPA);
            remera.setDescription("REMERA");
            remera.setQuantity(10);
            remera.setIsDeleted(false);
            remera.setLastDonationDate(today);
            remera.setMadeByOurselves(false);
            donations.add(remera);

            Donation chocolatada = new Donation();
            chocolatada.setCategory(Inventory.Category.ALIMENTOS);
            chocolatada.setDescription("CHOCOLATADA");
            chocolatada.setQuantity(7);
            chocolatada.setIsDeleted(false);
            chocolatada.setLastDonationDate(today);
            chocolatada.setMadeByOurselves(false);
            donations.add(chocolatada);

            Donation birome = new Donation();
            birome.setCategory(Inventory.Category.UTILES_ESCOLARES);
            birome.setDescription("BIROME");
            birome.setQuantity(14);
            birome.setIsDeleted(false);
            birome.setLastDonationDate(today);
            birome.setMadeByOurselves(false);
            donations.add(birome);

            Donation camioncito = new Donation();
            camioncito.setCategory(Inventory.Category.JUGUETES);
            camioncito.setDescription("CAMIONCITO");
            camioncito.setQuantity(3);
            camioncito.setIsDeleted(false);
            camioncito.setLastDonationDate(today);
            camioncito.setMadeByOurselves(false);
            donations.add(camioncito);

            Donation buzo = new Donation();
            buzo.setCategory(Inventory.Category.ROPA);
            buzo.setDescription("BUZO");
            buzo.setQuantity(8);
            buzo.setIsDeleted(false);
            buzo.setLastDonationDate(today.minusDays(1));
            buzo.setMadeByOurselves(true);
            donations.add(buzo);

            Donation fideos = new Donation();
            fideos.setCategory(Inventory.Category.ALIMENTOS);
            fideos.setDescription("FIDEOS");
            fideos.setQuantity(12);
            fideos.setIsDeleted(false);
            fideos.setLastDonationDate(today.minusDays(2));
            fideos.setMadeByOurselves(true);
            donations.add(fideos);

            Donation cuaderno = new Donation();
            cuaderno.setCategory(Inventory.Category.UTILES_ESCOLARES);
            cuaderno.setDescription("CUADERNO");
            cuaderno.setQuantity(20);
            cuaderno.setIsDeleted(false);
            cuaderno.setLastDonationDate(today.minusDays(3));
            cuaderno.setMadeByOurselves(false);
            donations.add(cuaderno);

            Donation pelota = new Donation();
            pelota.setCategory(Inventory.Category.JUGUETES);
            pelota.setDescription("PELOTA");
            pelota.setQuantity(5);
            pelota.setIsDeleted(false);
            pelota.setLastDonationDate(today.minusDays(4));
            pelota.setMadeByOurselves(false);
            donations.add(pelota);

            Donation gorro = new Donation();
            gorro.setCategory(Inventory.Category.ROPA);
            gorro.setDescription("GORRO");
            gorro.setQuantity(6);
            gorro.setIsDeleted(false);
            gorro.setLastDonationDate(today.minusDays(5));
            gorro.setMadeByOurselves(false);
            donations.add(gorro);

            Donation aceite = new Donation();
            aceite.setCategory(Inventory.Category.ALIMENTOS);
            aceite.setDescription("ACEITE");
            aceite.setQuantity(9);
            aceite.setIsDeleted(false);
            aceite.setLastDonationDate(today.minusDays(6));
            aceite.setMadeByOurselves(false);
            donations.add(aceite);

            Donation marcadores = new Donation();
            marcadores.setCategory(Inventory.Category.UTILES_ESCOLARES);
            marcadores.setDescription("MARCADORES");
            marcadores.setQuantity(15);
            marcadores.setIsDeleted(false);
            marcadores.setLastDonationDate(today.minusDays(7));
            marcadores.setMadeByOurselves(true);
            donations.add(marcadores);

            Donation lego = new Donation();
            lego.setCategory(Inventory.Category.JUGUETES);
            lego.setDescription("LEGO");
            lego.setQuantity(4);
            lego.setIsDeleted(false);
            lego.setLastDonationDate(today.minusDays(8));
            lego.setMadeByOurselves(true);
            donations.add(lego);

            donationService.saveAll(donations);
        }
    }

}
