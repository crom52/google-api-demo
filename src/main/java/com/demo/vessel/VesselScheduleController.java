package com.demo.vessel;

import java.util.List;
import lombok.RequiredArgsConstructor;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequiredArgsConstructor
class VesselScheduleController {
  private final VesselScheduleService vesselScheduleService;
  @GetMapping("vessel/schedule/next/eta-after-week")
  List getScheduleFromEmailAttachment() {
    return vesselScheduleService.findNext7DaysVesselPlan();
  }
}
