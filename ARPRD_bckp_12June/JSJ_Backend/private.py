import json
import math

def getJSON(filePathAndName):
    with open(filePathAndName, 'r') as fp:
        return json.load(fp)


def getpremiumprivate(Engin, VehicleVal, Cover, VehicleType, manload1, claimyears1, mandis, fleet, promotion, ct):
    vehicle = getJSON('./rating.json')
    EngineCC = int(Engin)
    VehicleVal1 = int(VehicleVal)
    fleetdisc = 0.12
    manload = float(manload1)
    claimyears = int(claimyears1)
    mandisc = float(mandis)

    #
    # def ct_prom(p4):
    #     prom = p4 * 0.083
    #     p5 = p4 - prom
    #     print("Annual Gross Premium with 12 for 11:" + str(p5))
    #     vot = p4 * 0.12
    #     p8 = vot
    #     p7 = p5 + p8
    #     print("Tax amount" + str(p8))
    #     print("Annual net premium " + str(p7))

    # def wct_prom(p4):
    #     prom = p4 * 0.083
    #     p5 = p4 - prom
    #     print("Annual Gross Premium with 12 for 11:" + str(p5))
    #     mot = p4 * 0.01
    #     p6 = p4 + mot
    #     vot = p6 * 0.12
    #     p7 = p5 + mot + vot
    #     p8 = mot + vot
    #     print("Tax amount" + str(p8))
    #     print("Annual net premium " + str(p7))

    def comp_fleet_base(base):

        if fleet:

            fle = 1 - fleetdisc
            mld = 1 + manload
            flm = fle * mld
            flm1 = flm - 1
            fpva = flm1 * 100
            inc = fpva / 100
            incr = 1 + inc
            p2 = base * incr

            return comp_ncd(p2)
        else:
            fle = 1 - 0
            mld = 1 + manload
            flm = fle * mld
            flm1 = flm - 1
            fpva = flm1 * 100
            inc = fpva / 100
            incr = 1 + inc
            p2 = base * incr

            return comp_ncd(p2)

    def comp_fleet_gp(gp):
        if fleet:
            fle = 1 - fleetdisc
            mld = 1 + manload
            flm = fle * mld
            flm1 = flm - 1
            fpva = flm1 * 100
            inc = fpva / 100
            incr = 1 + inc
            p2 = gp * incr

            return comp_ncd(p2)
        else:
            fle = 1 - 0
            mld = 1 + manload
            flm = fle * mld
            flm1 = flm - 1
            fpva = flm1 * 100
            inc = fpva / 100
            incr = 1 + inc
            p2 = gp * incr

            return comp_ncd(p2)

    def comp_fleet_vv(VV3):
        if fleet:
            fle = 1 - fleetdisc
            mld = 1 + manload
            flm = fle * mld
            flm1 = flm - 1
            fpva = flm1 * 100
            inc = fpva / 100
            incr = 1 + inc
            p2 = VV3 * incr

            return comp_ncd(p2)
        else:
            fle = 1 - 0
            mld = 1 + manload
            flm = fle * mld
            flm1 = flm - 1
            fpva = flm1 * 100
            inc = fpva / 100
            incr = 1 + inc
            p2 = VV3 * incr

            return comp_ncd(p2)

    def tp_fleet_vv(VV):
        if fleet:
            fle = 1 - fleetdisc
            mld = 1 + manload
            flm = fle * mld
            flm1 = flm - 1
            fpva = flm1 * 100
            inc = fpva / 100
            incr = 1 + inc
            p2 = float(VV)
            p1 = p2 * incr

            return tp_ncd2(p1)
        else:
            fle = 1 - 0
            mld = 1 + manload
            flm = fle * mld
            flm1 = flm - 1
            fpva = flm1 * 100
            inc = fpva / 100
            incr = 1 + inc
            p2 = float(VV)
            p1 = p2 * incr
            return tp_ncd2(p1)

    def tp_ncd2(p1):
        p4 = None
        p5 = None
        p7 = None
        p8 = None
        nc = None
        if claimyears == 0:
            ncd = vehicle['CY0'][0]["TP"]
            ncd2 = float(ncd)
            print("NCD Percentage:" + format(ncd2, ".0%"))
            nc = format(ncd2, ".0%")
            ncd1 = 1 - ncd2
            mandisc1 = 1
            manncd = ncd1 * mandisc1
            p3 = 1 - manncd
            bs = p1 * p3
            p4 = p1 - bs
            if p4 <= 300:
                p4 = 300
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif promotion:
                prom = p4 * 0.083
                p5 = p4 - prom
                print("Annual Gross Premium with 12 for 11:" + str(p5))
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p5 + mot + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
            else:
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p6 + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
        elif claimyears == 1:
            ncd = vehicle['CY1'][0]["TP"]
            ncd2 = float(ncd)
            print("NCD Percentage:" + format(ncd2, ".0%"))
            nc = format(ncd2, ".0%")
            ncd1 = 1 - ncd2
            mandisc1 = 1
            manncd = ncd1 * mandisc1
            p3 = 1 - manncd
            bs = p1 * p3
            p4 = p1 - bs
            if p4 <= 300:
                p4 = 300
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif promotion:
                prom = p4 * 0.083
                p5 = p4 - prom
                print("Annual Gross Premium with 12 for 11:" + str(p5))
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p5 + mot + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
            else:
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p6 + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
        elif claimyears == 2:
            ncd = vehicle['CY2'][0]["TP"]
            ncd2 = float(ncd)
            print("NCD Percentage:" + format(ncd2, ".0%"))
            nc = format(ncd2, ".0%")
            ncd1 = 1 - ncd2
            mandisc1 = 1
            manncd = ncd1 * mandisc1
            p3 = 1 - manncd
            bs = p1 * p3
            p4 = p1 - bs
            if p4 <= 300:
                p4 = 300
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif promotion:
                prom = p4 * 0.083
                p5 = p4 - prom
                print("Annual Gross Premium with 12 for 11:" + str(p5))
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p5 + mot + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
            else:
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p6 + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
        elif claimyears == 3:
            ncd = vehicle['CY3'][0]["TP"]
            ncd2 = float(ncd)
            print("NCD Percentage:" + format(ncd2, ".0%"))
            nc = format(ncd2, ".0%")
            ncd1 = 1 - ncd2
            mandisc1 = 1
            manncd = ncd1 * mandisc1
            p3 = 1 - manncd
            bs = p1 * p3
            p4 = p1 - bs
            if p4 <= 300:
                p4 = 300
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif promotion:
                prom = p4 * 0.083
                p5 = p4 - prom
                print("Annual Gross Premium with 12 for 11:" + str(p5))
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p5 + mot + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
            else:
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p6 + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
        elif claimyears == 4:

            ncd = vehicle['CY4'][0]["TP"]
            ncd2 = float(ncd)
            print("NCD Percentage:" + format(ncd2, ".0%"))
            nc = format(ncd2, ".0%")
            ncd1 = 1 - ncd2
            mandisc1 = 1
            manncd = ncd1 * mandisc1
            p3 = 1 - manncd
            bs = p1 * p3
            p4 = p1 - bs
            if p4 <= 300:
                p4 = 300
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif promotion:
                prom = p4 * 0.083
                p5 = p4 - prom
                print("Annual Gross Premium with 12 for 11:" + str(p5))
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p5 + mot + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
            else:
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p6 + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))

        elif claimyears == 5:
            ncd = vehicle['CY5'][0]["TP"]
            ncd2 = float(ncd)
            print("NCD Percentage:" + format(ncd2, ".0%"))
            nc = format(ncd2, ".0%")
            ncd1 = 1 - ncd2
            mandisc1 = 1
            manncd = ncd1 * mandisc1
            p3 = 1 - manncd
            bs = p1 * p3
            p4 = p1 - bs
            if p4 <= 300:
                p4 = 300
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif promotion:
                prom = p4 * 0.083
                p5 = p4 - prom
                print("Annual Gross Premium with 12 for 11:" + str(p5))
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p5 + mot + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
            else:
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p6 + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
        if promotion:
            p4 = round(p4, 2)
            p5 = round(p5, 2)
            p7 = round(p7, 2)
            p8 = round(p8, 2)
            return p4, p5, p7, p8, nc
        else:
            p4 = round(p4, 2)
            p7 = round(p7, 2)
            p8 = round(p8, 2)
            p5 = p4
            return p4, p5, p7, p8, nc


    def tpft_fleet_bp1(bp1):
        if fleet:
            fle = 1 - fleetdisc
            mld = 1 + manload
            flm = fle * mld
            flm1 = flm - 1
            fpva = flm1 * 100
            inc = fpva / 100
            incr = 1 + inc
            p2 = bp1 * incr

            return tpft_ncd(p2)
        else:
            fle = 1 - 0
            mld = 1 + manload
            flm = fle * mld
            flm1 = flm - 1
            fpva = flm1 * 100
            inc = fpva / 100
            incr = 1 + inc
            p2 = bp1 * incr

            return tpft_ncd(p2)

    def tpft_fleet_bp(bp):
        if fleet:
            fle = 1 - fleetdisc
            mld = 1 + manload
            flm = fle * mld
            flm1 = flm - 1
            fpva = flm1 * 100
            inc = fpva / 100
            incr = 1 + inc
            p2 = bp * incr

            return tpft_ncd(p2)
        else:
            fle = 1 - 0
            mld = 1 + manload
            flm = fle * mld
            flm1 = flm - 1
            fpva = flm1 * 100
            inc = fpva / 100
            incr = 1 + inc
            p2 = bp * incr

            return tpft_ncd(p2)

    def tpft_ncd(p2):
        p4 = None
        p5 = None
        p7 = None
        p8 = None
        nc = None
        if claimyears == 0:
            ncd = vehicle['CY0'][0]["TPFT"]
            ncd2 = float(ncd)
            print("NCD Percentage:" + format(ncd2, ".0%"))
            nc = format(ncd2, ".0%")
            ncd1 = 1 - ncd2
            mandisc1 = 1
            manncd = ncd1 * mandisc1
            p3 = 1 - manncd
            bs = p2 * p3
            p4 = p2 - bs
            if ct and p4 <= 500:
                p4 = 500
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif not ct and p4 <= 600:
                p4 = 600
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif promotion:
                prom = p4 * 0.083
                p5 = p4 - prom
                print("Annual Gross Premium with 12 for 11:" + str(p5))
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p5 + mot + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
            else:
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p6 + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
        elif claimyears == 1:
            ncd = vehicle['CY1'][0]["TPFT"]
            ncd2 = float(ncd)
            print("NCD Percentage:" + format(ncd2, ".0%"))
            nc = format(ncd2, ".0%")
            ncd1 = 1 - ncd2
            mandisc1 = 1
            manncd = ncd1 * mandisc1
            p3 = 1 - manncd
            bs = p2 * p3
            p4 = p2 - bs
            if ct and p4 <= 500:
                p4 = 500
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif not ct and p4 <= 600:
                p4 = 600
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif promotion:
                prom = p4 * 0.083
                p5 = p4 - prom
                print("Annual Gross Premium with 12 for 11:" + str(p5))
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p5 + mot + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
            else:
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p6 + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
        elif claimyears == 2:
            ncd = vehicle['CY2'][0]["TPFT"]
            ncd2 = float(ncd)
            print("NCD Percentage:" + format(ncd2, ".0%"))
            nc = format(ncd2, ".0%")
            ncd1 = 1 - ncd2
            mandisc1 = 1
            manncd = ncd1 * mandisc1
            p3 = 1 - manncd
            bs = p2 * p3
            p4 = p2 - bs
            if ct and p4 <= 500:
                p4 = 500
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif not ct and p4 <= 600:
                p4 = 600
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif promotion:
                prom = p4 * 0.083
                p5 = p4 - prom
                print("Annual Gross Premium with 12 for 11:" + str(p5))
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p5 + mot + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
            else:
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p6 + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
        elif claimyears == 3:
            ncd = vehicle['CY3'][0]["TPFT"]
            ncd2 = float(ncd)
            print("NCD Percentage:" + format(ncd2, ".0%"))
            nc = format(ncd2, ".0%")
            ncd1 = 1 - ncd2
            mandisc1 = 1
            manncd = ncd1 * mandisc1
            p3 = 1 - manncd
            bs = p2 * p3
            p4 = p2 - bs
            if ct and p4 <= 500:
                p4 = 500
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif not ct and p4 <= 600:
                p4 = 600
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif promotion:
                prom = p4 * 0.083
                p5 = p4 - prom
                print("Annual Gross Premium with 12 for 11:" + str(p5))
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p5 + mot + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
            else:
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p6 + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))

        elif claimyears == 4:
            ncd = vehicle['CY4'][0]["TPFT"]
            ncd2 = float(ncd)
            print("NCD Percentage:" + format(ncd2, ".0%"))
            nc = format(ncd2, ".0%")
            ncd1 = 1 - ncd2
            mandisc1 = 1
            manncd = ncd1 * mandisc1
            p3 = 1 - manncd
            bs = p2 * p3
            p4 = p2 - bs
            if ct and p4 <= 500:
                p4 = 500
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif not ct and p4 <= 600:
                p4 = 600
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif promotion:
                prom = p4 * 0.083
                p5 = p4 - prom
                print("Annual Gross Premium with 12 for 11:" + str(p5))
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p5 + mot + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
            else:
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p6 + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))

        elif claimyears == 5:
            ncd = vehicle['CY5'][0]["TPFT"]
            ncd2 = float(ncd)
            print("NCD Percentage:" + format(ncd2, ".0%"))
            nc = format(ncd2, ".0%")
            ncd1 = 1 - ncd2
            mandisc1 = 1
            manncd = ncd1 * mandisc1
            p3 = 1 - manncd
            bs = p2 * p3
            p4 = p2 - bs
            if ct and p4 <= 500:
                p4 = 500
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif not ct and p4 <= 600:
                p4 = 600
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif promotion:
                prom = p4 * 0.083
                p5 = p4 - prom
                print("Annual Gross Premium with 12 for 11:" + str(p5))
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p5 + mot + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
            else:
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p6 + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
        if promotion:
            p4 = round(p4, 2)
            p5 = round(p5, 2)
            p7 = round(p7, 2)
            p8 = round(p8, 2)
            return p4, p5, p7, p8, nc
        else:
            p4 = round(p4, 2)
            p7 = round(p7, 2)
            p8 = round(p8, 2)
            p5 = p4
            return p4, p5, p7, p8, nc
    def comp_ncd(p2):
        p4 = None
        p5 = None
        p7 = None
        p8 = None
        nc = None
        if claimyears == 0:
            ncd = vehicle['CY0'][0]["COMP"]
            ncd2 = float(ncd)
            print("NCD Percentage:" + format(ncd2, ".0%"))
            nc = format(ncd2, ".0%")
            ncd1 = 1 - ncd2
            mandisc1 = 1 - mandisc
            manncd = ncd1 * mandisc1
            p3 = 1 - manncd
            bs = p2 * p3
            p4 = p2 - bs
            if p4 <= 550:
                p4 = 550
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif promotion:
                prom = p4 * 0.083
                p5 = p4 - prom
                print("Annual Gross Premium with 12 for 11:" + str(p5))
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p5 + mot + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
            else:
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p6 + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
        elif claimyears == 1:
            ncd = vehicle['CY1'][0]["COMP"]
            ncd2 = float(ncd)
            print("NCD Percentage:" + format(ncd2, ".0%"))
            nc = format(ncd2, ".0%")
            ncd1 = 1 - ncd2
            mandisc1 = 1 - mandisc
            manncd = ncd1 * mandisc1
            p3 = 1 - manncd
            bs = p2 * p3
            p4 = p2 - bs
            if p4 <= 550:
                p4 = 550
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif promotion:
                prom = p4 * 0.083
                p5 = p4 - prom
                print("Annual Gross Premium with 12 for 11:" + str(p5))
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p5 + mot + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
            else:
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p6 + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
        elif claimyears == 2:
            ncd = vehicle['CY2'][0]["COMP"]
            ncd2 = float(ncd)
            print("NCD Percentage:" + format(ncd2, ".0%"))
            nc = format(ncd2, ".0%")
            ncd1 = 1 - ncd2
            mandisc1 = 1 - mandisc
            manncd = ncd1 * mandisc1
            p3 = 1 - manncd
            bs = p2 * p3
            p4 = p2 - bs
            if p4 <= 550:
                p4 = 550
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif promotion:
                prom = p4 * 0.083
                p5 = p4 - prom
                print("Annual Gross Premium with 12 for 11:" + str(p5))
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p5 + mot + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
            else:
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p6 + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
        elif claimyears == 3:
            ncd = vehicle['CY3'][0]["COMP"]
            ncd2 = float(ncd)
            print("NCD Percentage:" + format(ncd2, ".0%"))
            nc = format(ncd2, ".0%")
            ncd1 = 1 - ncd2
            mandisc1 = 1 - mandisc
            manncd = ncd1 * mandisc1
            p3 = 1 - manncd
            bs = p2 * p3
            p4 = p2 - bs
            if p4 <= 550:
                p4 = 550
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif promotion:
                prom = p4 * 0.083
                p5 = p4 - prom
                print("Annual Gross Premium with 12 for 11:" + str(p5))
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p5 + mot + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
            else:
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p6 + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
        elif claimyears == 4:
            ncd = vehicle['CY4'][0]["COMP"]
            ncd2 = float(ncd)
            print("NCD Percentage:" + format(ncd2, ".0%"))
            nc = format(ncd2, ".0%")
            ncd1 = 1 - ncd2
            mandisc1 = 1 - mandisc
            manncd = ncd1 * mandisc1
            p3 = 1 - manncd
            bs = p2 * p3
            p4 = p2 - bs
            if p4 <= 550:
                p4 = 550
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif promotion:
                prom = p4 * 0.083
                p5 = p4 - prom
                print("Annual Gross Premium with 12 for 11:" + str(p5))
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p5 + mot + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
            else:
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p6 + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
        elif claimyears == 5:
            ncd = vehicle['CY5'][0]["COMP"]
            ncd2 = float(ncd)
            print("NCD Percentage:" + format(ncd2, ".0%"))
            nc = format(ncd2, ".0%")
            ncd1 = 1 - ncd2
            mandisc1 = 1 - mandisc
            manncd = ncd1 * mandisc1
            p3 = 1 - manncd
            bs = p2 * p3
            p4 = p2 - bs
            if p4 <= 550:
                p4 = 550
                print("Minimum premium:  " + str(p4))
                if promotion:
                    prom = p4 * 0.083
                    p5 = p4 - prom
                    print("Annual Gross Premium with 12 for 11:" + str(p5))
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p5 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
                else:
                    vot = p4 * 0.12
                    p8 = vot
                    p7 = p4 + p8
                    print("Tax amount" + str(p8))
                    print("Annual net premium " + str(p7))
            elif promotion:
                prom = p4 * 0.083
                p5 = p4 - prom
                print("Annual Gross Premium with 12 for 11:" + str(p5))
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p5 + mot + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
            else:
                mot = p4 * 0.01
                p6 = p4 + mot
                vot = p6 * 0.12
                p7 = p6 + vot
                p8 = mot + vot
                print("Tax amount" + str(p8))
                print("Annual net premium " + str(p7))
        if promotion:
            p4 = round(p4, 2)
            p5 = round(p5, 2)
            p7 = round(p7, 2)
            p8 = round(p8, 2)
            return p4, p5, p7, p8, nc
        else:
            p4 = round(p4, 2)
            p7 = round(p7, 2)
            p8 = round(p8, 2)
            p5 = p4
            return p4, p5, p7, p8, nc

    if EngineCC in range(0, 1400) and Cover == "COMP":
        EngineCC = 1400
    elif EngineCC in range(1401, 2200) and Cover == "COMP":
        EngineCC = 2200
    elif EngineCC in range(2201, 3050) and Cover == "COMP":
        EngineCC = 3050
    elif EngineCC in range(3051, 3650) and Cover == "COMP":
        EngineCC = 3650
    elif EngineCC in range(3651, 4400) and Cover == "COMP":
        EngineCC = 4400
    elif EngineCC > 4400 and Cover == "COMP":
        EngineCC = '>4400'
    elif EngineCC in range(0, 2200) and Cover == "TP":
        EngineCC = 2200
    elif EngineCC in range(2201, 4400) and Cover == "TP":
        EngineCC = 4400
    elif EngineCC in range(4401, 5800) and Cover == "TP":
        EngineCC = 5800
    elif EngineCC > 5800 and Cover == "TP":
        EngineCC = '>5800'
    elif EngineCC in range(0, 2200) and Cover == "TPFT":
        EngineCC = 2200
    elif EngineCC in range(2201, 4400) and Cover == "TPFT":
        EngineCC = 4400
    elif EngineCC in range(4401, 5800) and Cover == "TPFT":
        EngineCC = 5800
    elif EngineCC > 5800 and Cover == "TPFT":
        EngineCC = '>5800'
    if VehicleVal1 in range(0, 3500) and Cover == "COMP":
        VehicleVal1 = 3500
    elif VehicleVal1 in range(3501, 4000) and Cover == "COMP":
        VehicleVal1 = 4000
    elif VehicleVal1 in range(4001, 4500) and Cover == "COMP":
        VehicleVal1 = 4500
    elif VehicleVal1 in range(4501, 5000) and Cover == "COMP":
        VehicleVal1 = 5000
    elif VehicleVal1 in range(5001, 5500) and Cover == "COMP":
        VehicleVal1 = 5500
    elif VehicleVal1 in range(5501, 6000) and Cover == "COMP":
        VehicleVal1 = 6000
    elif VehicleVal1 in range(6001, 6500) and Cover == "COMP":
        VehicleVal1 = 6500
    elif VehicleVal1 in range(6501, 7000) and Cover == "COMP":
        VehicleVal1 = 7000
    elif VehicleVal1 in range(7001, 7500) and Cover == "COMP":
        VehicleVal1 = 7500
    elif VehicleVal1 in range(7501, 8000) and Cover == "COMP":
        VehicleVal1 = 8000
    elif VehicleVal1 in range(8001, 8500) and Cover == "COMP":
        VehicleVal1 = 8500
    elif VehicleVal1 in range(8501, 9000) and Cover == "COMP":
        VehicleVal1 = 9000
    elif VehicleVal1 in range(9001, 9500) and Cover == "COMP":
        VehicleVal1 = 9500
    elif VehicleVal1 in range(9501, 10000) and Cover == "COMP":
        VehicleVal1 = 10000
    if EngineCC == 1400 and Cover == "COMP":
        if VehicleVal1 >= 10000:
            VV1 = 10000
            if VehicleType == "Golf":
                VV2 = vehicle['Set1'][0]['VehicleValue'][0][str(VV1)]
                CC = vehicle['Set1'][0]['CubicCentimeters']
                CI = vehicle['Set1'][0]['CubicInches']
                VV3 = int(VV2)
                bp = VehicleVal1 - 10000
                bp1 = bp * 0.075

                base = VV3 + bp1
                print("Comprehensive Base premium: " + str(base))
                gp = base * 0.75
                print("Golf cart comprehensive base premium: " + str(gp))
                return comp_fleet_gp(gp)
            elif VehicleType == "Electric":
                VV1 = 10000
                VV2 = vehicle['Set2'][0]['VehicleValue'][0][str(VV1)]
                CC = vehicle['Set2'][0]['CubicCentimeters']
                CI = vehicle['Set2'][0]['CubicInches']
                VV3 = int(VV2)
                bp = VehicleVal1 - 10000
                bp1 = bp * 0.075

                base = VV3 + bp1
                print("Electric comprehensive base premium: " + str(base))
                return comp_fleet_base(base)
            elif VehicleType == "NoType":
                VV2 = vehicle['Set1'][0]['VehicleValue'][0][str(VV1)]
                CC = vehicle['Set1'][0]['CubicCentimeters']
                CI = vehicle['Set1'][0]['CubicInches']
                VV3 = int(VV2)
                bp = VehicleVal1 - 10000
                bp1 = bp * 0.075

                base = VV3 + bp1
                print("Comprehensive Base premium: " + str(base))
                return comp_fleet_base(base)
            else:
                print("Error in 1400 comp 10000")

        else:
            if VehicleType == "Golf":
                VV2 = vehicle['Set1'][0]['VehicleValue'][0][str(VehicleVal1)]
                CC = vehicle['Set1'][0]['CubicCentimeters']
                CI = vehicle['Set1'][0]['CubicInches']
                VV3 = int(VV2)
                gp = VV3 * 0.75
                print("Golf cart comprehensive base premium: " + str(gp))
                return comp_fleet_gp(gp)
            elif VehicleType == "Electric" and EngineCC <= 2200:
                VV2 = vehicle['Set2'][0]['VehicleValue'][0][str(VehicleVal1)]
                CC = vehicle['Set2'][0]['CubicCentimeters']
                CI = vehicle['Set2'][0]['CubicInches']
                VV3 = int(VV2)
                base = VV3
                print("Electric comprehensive base premium: " + str(base))
                return comp_fleet_base(base)
            elif VehicleType == "Electric" and EngineCC > 2200:
                VV2 = vehicle['Set2'][0]['VehicleValue'][0][str(VehicleVal1)]
                CC = vehicle['Set2'][0]['CubicCentimeters']
                CI = vehicle['Set2'][0]['CubicInches']
                VV3 = int(VV2)
                base = VV3
                print("Electric comprehensive base premium: " + str(base))
                return comp_fleet_base(base)
            elif VehicleType == "NoType":
                VV2 = vehicle['Set1'][0]['VehicleValue'][0][str(VehicleVal1)]
                CC = vehicle['Set1'][0]['CubicCentimeters']
                CI = vehicle['Set1'][0]['CubicInches']
                print(VV2)
                VV3 = int(VV2)
                print("Normal Comprehensive Base premium: " + str(VV3))
                return comp_fleet_vv(VV3)
            else:
                print("Error in 1400 comp <10000")
    elif EngineCC == 2200 and Cover == "COMP":
        if VehicleVal1 >= 10000:
            VV1 = 10000
            if VehicleType == "Golf":
                VV2 = vehicle['Set2'][0]['VehicleValue'][0][str(VV1)]
                CC = vehicle['Set2'][0]['CubicCentimeters']
                CI = vehicle['Set2'][0]['CubicInches']
                VV3 = int(VV2)
                bp = VehicleVal1 - 10000
                bp1 = bp * 0.075

                base = VV3 + bp1
                print("Comprehensive Base premium: " + str(base))
                gp = base * 0.75
                print("Golf cart comprehensive base premium: " + str(gp))
                return comp_fleet_gp(gp)
            elif VehicleType == "Electric":
                VV1 = 10000
                VV2 = vehicle['Set2'][0]['VehicleValue'][0][str(VV1)]
                CC = vehicle['Set2'][0]['CubicCentimeters']
                CI = vehicle['Set2'][0]['CubicInches']
                VV3 = int(VV2)
                bp = VehicleVal1 - 10000
                bp1 = bp * 0.075

                base = VV3 + bp1
                print("Electric comprehensive base premium: " + str(base))
                return comp_fleet_base(base)
            elif VehicleType == "NoType":
                VV2 = vehicle['Set2'][0]['VehicleValue'][0][str(VV1)]
                CC = vehicle['Set2'][0]['CubicCentimeters']
                CI = vehicle['Set2'][0]['CubicInches']
                VV3 = int(VV2)
                bp = VehicleVal1 - 10000
                bp1 = bp * 0.075

                base = VV3 + bp1
                print("Comprehensive Base premium: " + str(base))
                return comp_fleet_base(base)
            else:
                print("Error in 2200 comp 10000")
        else:

            if VehicleType == "Golf":
                VV2 = vehicle['Set2'][0]['VehicleValue'][0][str(VehicleVal1)]
                CC = vehicle['Set2'][0]['CubicCentimeters']
                CI = vehicle['Set2'][0]['CubicInches']
                VV3 = int(VV2)
                gp = VV3 * 0.75
                print("Golf cart comprehensive base premium: " + str(gp))
                return comp_fleet_gp(gp)
            elif VehicleType == "Electric":
                VV2 = vehicle['Set2'][0]['VehicleValue'][0][str(VehicleVal1)]
                CC = vehicle['Set2'][0]['CubicCentimeters']
                CI = vehicle['Set2'][0]['CubicInches']
                VV3 = int(VV2)
                base = VV3
                print("Electric comprehensive base premium: " + str(base))
                return comp_fleet_base(base)
            elif VehicleType == "NoType":
                VV2 = vehicle['Set2'][0]['VehicleValue'][0][str(VehicleVal1)]
                CC = vehicle['Set2'][0]['CubicCentimeters']
                CI = vehicle['Set2'][0]['CubicInches']
                VV3 = int(VV2)
                base = VV3
                print("Normal Comprehensive Base premium: " + str(base))
                return comp_fleet_base(base)
            else:
                print("Error in 2200 comp <10000")

    elif EngineCC == 3050 and Cover == "COMP":
        if VehicleVal1 >= 10000:
            VV1 = 10000
            if VehicleType == "Golf":
                VV2 = vehicle['Set3'][0]['VehicleValue'][0][str(VV1)]
                CC = vehicle['Set3'][0]['CubicCentimeters']
                CI = vehicle['Set3'][0]['CubicInches']
                VV3 = int(VV2)
                bp = VehicleVal1 - 10000
                bp1 = bp * 0.075

                base = VV3 + bp1
                print("Comprehensive Base premium: " + str(base))
                gp = base * 0.75
                print("Golf cart comprehensive base premium: " + str(gp))
                return comp_fleet_gp(gp)
            elif VehicleType == "NoType" or VehicleType == "Electric":
                VV2 = vehicle['Set3'][0]['VehicleValue'][0][str(VV1)]
                CC = vehicle['Set3'][0]['CubicCentimeters']
                CI = vehicle['Set3'][0]['CubicInches']
                VV3 = int(VV2)
                bp = VehicleVal1 - 10000
                bp1 = bp * 0.075

                base = VV3 + bp1
                print("Comprehensive Base premium: " + str(base))
                return comp_fleet_base(base)
            else:
                print("Error in 3050 comp 10000")
        else:

            if VehicleType == "Golf":
                VV2 = vehicle['Set3'][0]['VehicleValue'][0][str(VehicleVal1)]
                CC = vehicle['Set3'][0]['CubicCentimeters']
                CI = vehicle['Set3'][0]['CubicInches']
                VV3 = int(VV2)
                gp = VV3 * 0.75
                print("Golf cart comprehensive base premium: " + str(gp))
                return comp_fleet_gp(gp)
            elif VehicleType == "NoType" or VehicleType == "Electric":
                VV2 = vehicle['Set3'][0]['VehicleValue'][0][str(VehicleVal1)]
                CC = vehicle['Set3'][0]['CubicCentimeters']
                CI = vehicle['Set3'][0]['CubicInches']
                VV3 = int(VV2)
                base = VV3
                print("Normal Comprehensive Base premium: " + str(base))
                return comp_fleet_base(base)
            else:
                print("Error in 3050 comp <10000")

    elif EngineCC == 3650 and Cover == "COMP":
        if VehicleVal1 >= 10000:
            VV1 = 10000
            if VehicleType == "Golf":
                VV2 = vehicle['Set4'][0]['VehicleValue'][0][str(VV1)]
                CC = vehicle['Set4'][0]['CubicCentimeters']
                CI = vehicle['Set4'][0]['CubicInches']
                VV3 = int(VV2)
                bp = VehicleVal1 - 10000
                bp1 = bp * 0.075

                base = VV3 + bp1
                print("Comprehensive Base premium: " + str(base))
                gp = base * 0.75
                print("Golf cart comprehensive base premium: " + str(gp))
                return comp_fleet_gp(gp)
            elif VehicleType == "NoType" or VehicleType == "Electric":
                VV2 = vehicle['Set4'][0]['VehicleValue'][0][str(VV1)]
                CC = vehicle['Set4'][0]['CubicCentimeters']
                CI = vehicle['Set4'][0]['CubicInches']
                VV3 = int(VV2)
                bp = VehicleVal1 - 10000
                bp1 = bp * 0.075

                base = VV3 + bp1
                print("Comprehensive Base premium: " + str(base))
                return comp_fleet_base(base)
            else:
                print("Error in 3650 comp 10000")
        else:
            if VehicleType == "Golf":
                VV2 = vehicle['Set4'][0]['VehicleValue'][0][str(VehicleVal1)]
                CC = vehicle['Set4'][0]['CubicCentimeters']
                CI = vehicle['Set4'][0]['CubicInches']
                VV3 = int(VV2)
                gp = VV3 * 0.75
                print("Golf cart comprehensive base premium: " + str(gp))
                return comp_fleet_gp(gp)
            elif VehicleType == "NoType" or VehicleType == "Electric":
                VV2 = vehicle['Set4'][0]['VehicleValue'][0][str(VehicleVal1)]
                CC = vehicle['Set4'][0]['CubicCentimeters']
                CI = vehicle['Set4'][0]['CubicInches']
                print(VV2)
                VV3 = int(VV2)
                base = VV3
                print("Normal Comprehensive Base premium: " + str(base))
                return comp_fleet_base(base)
            else:
                print("Error in 3650 comp <10000")

    elif EngineCC == 4400 and Cover == "COMP":
        if VehicleVal1 >= 10000:
            VV1 = 10000
            if VehicleType == "Golf":
                VV2 = vehicle['Set5'][0]['VehicleValue'][0][str(VV1)]
                CC = vehicle['Set5'][0]['CubicCentimeters']
                CI = vehicle['Set5'][0]['CubicInches']
                VV3 = int(VV2)
                bp = VehicleVal1 - 10000
                bp1 = bp * 0.075

                base = VV3 + bp1
                print("Comprehensive Base premium: " + str(base))
                gp = base * 0.75
                print("Golf cart comprehensive base premium: " + str(gp))
                return comp_fleet_gp(gp)
            elif VehicleType == "NoType" or VehicleType == "Electric":
                VV2 = vehicle['Set5'][0]['VehicleValue'][0][str(VV1)]
                CC = vehicle['Set5'][0]['CubicCentimeters']
                CI = vehicle['Set5'][0]['CubicInches']
                VV3 = int(VV2)
                bp = VehicleVal1 - 10000
                bp1 = bp * 0.075

                base = VV3 + bp1
                print("Comprehensive Base premium: " + str(base))
                return comp_fleet_base(base)
            else:
                print("Error in 4400 comp 10000")
        else:

            if VehicleType == "Golf":
                VV2 = vehicle['Set5'][0]['VehicleValue'][0][str(VehicleVal1)]
                CC = vehicle['Set5'][0]['CubicCentimeters']
                CI = vehicle['Set5'][0]['CubicInches']
                VV3 = int(VV2)
                gp = VV3 * 0.75
                print("Golf cart comprehensive base premium: " + str(gp))
                return comp_fleet_gp(gp)
            elif VehicleType == "NoType" or VehicleType == "Electric":
                VV2 = vehicle['Set5'][0]['VehicleValue'][0][str(VehicleVal1)]
                CC = vehicle['Set5'][0]['CubicCentimeters']
                CI = vehicle['Set5'][0]['CubicInches']
                print(VV2)
                VV3 = int(VV2)
                base = VV3
                print("Normal Comprehensive Base premium: " + str(base))
                return comp_fleet_base(base)
            else:
                print("Error in 4400 comp <10000")

    elif EngineCC == '>4400' and Cover == "COMP":
        if VehicleVal1 >= 10000:
            VV1 = 10000
            if VehicleType == "Golf":
                VV2 = vehicle['Set6'][0]['VehicleValue'][0][str(VV1)]
                CC = vehicle['Set6'][0]['CubicCentimeters']
                CI = vehicle['Set6'][0]['CubicInches']
                VV3 = int(VV2)
                bp = VehicleVal1 - 10000
                bp1 = bp * 0.075

                base = VV3 + bp1
                print("Comprehensive Base premium: " + str(base))
                gp = base * 0.75
                print("Golf cart comprehensive base premium: " + str(gp))
                return comp_fleet_gp(gp)
            elif VehicleType == "NoType" or VehicleType == "Electric":
                VV2 = vehicle['Set6'][0]['VehicleValue'][0][str(VV1)]
                CC = vehicle['Set6'][0]['CubicCentimeters']
                CI = vehicle['Set6'][0]['CubicInches']
                VV3 = int(VV2)
                bp = VehicleVal1 - 10000
                bp1 = bp * 0.075

                base = VV3 + bp1
                print("Comprehensive Base premium: " + str(base))
                return comp_fleet_base(base)
            else:
                print("Error in >4400 comp 10000")
        else:
            if VehicleType == "Golf":
                VV2 = vehicle['Set6'][0]['VehicleValue'][0][str(VehicleVal1)]
                CC = vehicle['Set6'][0]['CubicCentimeters']
                CI = vehicle['Set6'][0]['CubicInches']
                VV3 = int(VV2)
                gp = VV3 * 0.75
                print("Golf cart comprehensive base premium: " + str(gp))
                return comp_fleet_gp(gp)
            elif VehicleType == "NoType" or VehicleType == "Electric":
                VV2 = vehicle['Set6'][0]['VehicleValue'][0][str(VehicleVal1)]
                CC = vehicle['Set6'][0]['CubicCentimeters']
                CI = vehicle['Set6'][0]['CubicInches']
                print(VV2)
                VV3 = int(VV2)
                base = VV3
                print("Normal Comprehensive Base premium: " + str(base))
                return comp_fleet_base(base)
            else:
                print("Error >4400 in comp <10000")


    elif EngineCC == 2200 and Cover == "TP":
        if VehicleType == "Golf":
            VV = vehicle['TP'][0]['2200']
            print("Base premium Third Party: " + VV)
            VV1 = int(VV)
            VV = VV1 * 0.75
            print("Golf cart Third Party base premium: " + str(VV))
            return tp_fleet_vv(VV)
        elif VehicleType == "Electric":
            VV = vehicle['TP'][0]['2200']
            print("Electric comprehensive base premium: " + VV)
            return tp_fleet_vv(VV)
        elif VehicleType == "NoType":
            VV = vehicle['TP'][0]['2200']
            print("Base premium Third Party: " + VV)
            return tp_fleet_vv(VV)
        else:
            print("Error in TP 2200")
    elif EngineCC == 4400 and Cover == "TP":
        if VehicleType == "Golf":
            VV = vehicle['TP'][0]['4400']
            print("Base premium Third Party: " + VV)
            VV1 = int(VV)
            VV = VV1 * 0.75
            print("Golf cart Third Party base premium: " + str(VV))
            return tp_fleet_vv(VV)
        elif VehicleType == "NoType" or VehicleType == "Electric":
            VV = vehicle['TP'][0]['4400']
            print("Normal Base premium Third Party: " + VV)
            return tp_fleet_vv(VV)
        else:
            print("Error in 4400 TP")
    elif EngineCC == 5800 and Cover == "TP":
        if VehicleType == "Golf":
            VV = vehicle['TP'][0]['5800']
            print("Base premium Third Party: " + VV)
            VV1 = int(VV)
            VV = VV1 * 0.75
            print("Golf cart Third Party base premium: " + str(VV))
            return tp_fleet_vv(VV)
        elif VehicleType == "NoType" or VehicleType == "Electric":
            VV = vehicle['TP'][0]['5800']
            print("Base premium Third Party: " + VV)
            return tp_fleet_vv(VV)
        else:
            print("Error in 5800 TP")
    elif EngineCC == '>5800' and Cover == "TP":
        if VehicleType == "Golf":
            VV = vehicle['TP'][0]['>5800']
            print("Base premium Third Party: " + VV)
            VV1 = int(VV)
            VV = VV1 * 0.75
            print("Golf cart Third Party base premium: " + str(VV))
            return tp_fleet_vv(VV)
        elif VehicleType == "NoType" or VehicleType == "Electric":
            VV = vehicle['TP'][0]['>5800']
            print("Normal Base premium Third Party: " + VV)
            return tp_fleet_vv(VV)
        else:
            print("Error in >5800 TP")

    elif EngineCC == 2200 and Cover == "TPFT":
        if VehicleType == "Golf":
            VV = vehicle['TPFT'][0]['2200']
            VV1 = int(VV)
            print(VV1)
            Vehval1 = VehicleVal1 * 0.04
            bp = VV1 + Vehval1
            bp1 = bp * 0.75
            print("Base premium Third Party Fire & Theft Golf" + str(bp1))
            return tpft_fleet_bp1(bp1)
        elif VehicleType == "NoType" or VehicleType == "Electric":
            VV = vehicle['TPFT'][0]['2200']
            VV1 = int(VV)
            VV2 = 0.04 * VehicleVal1
            bp = VV1 + VV2
            print("TPFT Base premium of notype: " + str(bp))
            return tpft_fleet_bp(bp)
        else:
            print("Error in TPFT 2200")
    elif EngineCC == 4400 and Cover == "TPFT":
        if VehicleType == "Golf":
            VV = vehicle['TPFT'][0]['4400']
            VV1 = int(VV)
            print(VV1)
            Vehval1 = VehicleVal1 * 0.04
            bp = VV1 + Vehval1
            bp1 = bp * 0.75
            print("Base premium Third Party Fire & Theft Golf" + str(bp1))
            return tpft_fleet_bp1(bp1)
        elif VehicleType == "NoType" or VehicleType == "Electric":
            VV = vehicle['TPFT'][0]['4400']
            VV1 = int(VV)
            VV2 = 0.04 * VehicleVal1
            bp = VV1 + VV2
            print("TPFT Base premium of notype: " + str(bp))
            return tpft_fleet_bp(bp)
        else:
            print("Error in TPFT 4400")
    elif EngineCC == 5800 and Cover == "TPFT":
        if VehicleType == "Golf":
            VV = vehicle['TPFT'][0]['5800']
            VV1 = int(VV)
            print(VV1)
            Vehval1 = VehicleVal1 * 0.04
            bp = VV1 + Vehval1
            bp1 = bp * 0.75
            print("Base premium Third Party Fire & Theft Golf" + str(bp1))
            return tpft_fleet_bp1(bp1)
        elif VehicleType == "NoType" or VehicleType == "Electric":
            VV = vehicle['TPFT'][0]['5800']
            VV1 = int(VV)
            VV2 = 0.04 * VehicleVal1
            bp = VV1 + VV2
            print("TPFT Base premium of notype: " + str(bp))
            return tpft_fleet_bp(bp)
        else:
            print("Error in TPFT 5800")
    elif EngineCC == '>5800' and Cover == "TPFT":
        if VehicleType == "Golf":
            VV = vehicle['TPFT'][0]['>5800']
            VV1 = int(VV)
            print(VV1)
            Vehval1 = VehicleVal1 * 0.04
            bp = VV1 + Vehval1
            bp1 = bp * 0.75
            print("Base premium Third Party Fire & Theft Golf" + str(bp1))
            return tpft_fleet_bp1(bp1)
        elif VehicleType == "NoType" or VehicleType == "Electric":
            VV = vehicle['TPFT'][0]['>5800']
            VV1 = int(VV)
            VV2 = 0.04 * VehicleVal1
            bp = VV1 + VV2
            print("TPFT Base premium of notype: " + str(bp))
            return tpft_fleet_bp(bp)
        else:
            print("Error in TPFT >5800")
    else:
        print("No value entered")
print(getpremiumprivate(1400, 5200, 'TPFT', 'NoType', 0, 5, 0, False, False, True))