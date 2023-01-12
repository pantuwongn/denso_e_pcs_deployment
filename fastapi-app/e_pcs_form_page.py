class PCSPageBuilder:
    def __init__(self, maxHeight, pcsFormDict: dict):
        self.maxHeight = maxHeight
        self.pcsFormDict = pcsFormDict
    
    def create(self):
        i = 1
        currentPageLineCount = 0
        pcsPageItemDict = dict()

        for processItem in self.pcsFormDict['items']:
            pcsItemPage = PCSItemPage(processItem)
            currentPageLineCount += pcsItemPage.height

            if currentPageLineCount > self.maxHeight:
                i += 1
                currentPageLineCount = pcsItemPage.height
            
            if pcsPageItemDict.get(i, None) is None:
                pcsPageItemDict[i] = [pcsItemPage]
            else:
                pcsPageItemDict[i].append(pcsItemPage)

        pcsPageList = list()
        pcsPageItemEnumerable = pcsPageItemDict.items()
        totalPage = len(pcsPageItemEnumerable)
        for page, pcsPageItemList in pcsPageItemEnumerable:
            pcsPageList.append(PCSHeaderPage(page, totalPage, self.pcsFormDict, pcsPageItemList))
        
        return pcsPageList

class PCSHeaderPage:
    def __init__(self, page: int, totalPage: int, pcsFormDict: dict, pcsItemList: list):
        self.page = page
        self.totalPage = totalPage
        self.pcsFormDict = pcsFormDict
        self.pcsItemList = pcsItemList

    @property
    def lineName(self):
        return self.pcsFormDict['line']
    
    @property
    def assyName(self):
        return self.pcsFormDict['assy_name']

    @property
    def processName(self):
        return self.pcsFormDict['process_name']
    
    @property
    def partName(self):
        return self.pcsFormDict['part_name']
    
    @property
    def customerName(self):
        return self.pcsFormDict['customer']

    @property
    def itemList(self):
        return self.pcsItemList

    @property
    def pageCount(self):
        return '{} / {}'.format(self.page, self.totalPage)

    @property
    def criticalItemSummaryDict(self):
        criticalItemDict = dict()

        for item in self.itemList:
            scSymbolList = item.scSymbolList
            for scSymbol in scSymbolList:
                if scSymbol['character'] == 'RW':
                    continue
                symbolHash = '{}-{}'.format(scSymbol['character'], scSymbol['shape'])

                if criticalItemDict.get(symbolHash, None) is None:
                    criticalItemDict[symbolHash] = 1
                else:
                    criticalItemDict[symbolHash] += 1
        return criticalItemDict

    @property
    def summaryTimingDict(self):
        duringList = list()
        beforeList = list()
        afterList = list()

        for i, pcsItem in enumerate(self.pcsItemList):
            if pcsItem.checkTiming == 'During':
                duringList.append(i)
            elif pcsItem.checkTiming == 'Before':
                beforeList.append(i)
            elif pcsItem.checkTiming == 'After':
                afterList.append(i)

        return {
            'During': duringList,
            'Before': beforeList,
            'After': afterList
        }

class PCSItemPage:
    def __init__(self, processItem):
        self.processItem = processItem

    @property
    def height(self):
        def _parameterColHeight():
            return len(self.parameterDetail) + len(self.measuredValue) + len(self.measurement)

        def _controlMethodIntervalColHeight():
            return len(self.controlMethodIntervalPeriod) + len(self.controlMethodCalibrationInterval) + len(self.controlMethodSampleNo) + len(self.controlMethodIntervalType)

        def _controlMethodColHeight():
            return 1 + len(self.controlMethodType) + len(self.controlMethodItemType)

        def _controlMethodControlPersonColHeight():
            return 1 + len(self.controlMethodControlPerson)
            
        def _processCapabilityColHeight():
            return len(self.processCapability)

        def _remarkHeightColHeight():
            return len(self.remark)

        def _relateStandardColHeight():
            return len(self.relateStandard)
        
        maxHeight = max(
            _parameterColHeight(),
            _controlMethodIntervalColHeight(),
            _controlMethodColHeight(),
            _controlMethodControlPersonColHeight(),
            _processCapabilityColHeight(),
            _remarkHeightColHeight(),
            _relateStandardColHeight()
        )
        return 3 if maxHeight < 3 else maxHeight

    @property
    def scSymbolList(self):
        return self.processItem['sc_symbols']

    @property
    def parameterDetail(self):
        return self._splitLine(self.processItem['parameter']['parameter'])

    @property
    def measuredValue(self):
        def _appendTextIfExist(targetStr: str, dataDict: dict, key: str, spaced = False):
            finalText = targetStr
            if dataDict.get(key, '').strip() != '':
                finalText = '{}{}{}'.format(targetStr, ' ' if spaced else '', dataDict[key])
            return finalText

        result = ''
        result = _appendTextIfExist(result, self.processItem['parameter'], 'prefix')
        result = _appendTextIfExist(result, self.processItem['parameter'], 'main')
        result = _appendTextIfExist(result, self.processItem['parameter'], 'suffix')
        result = _appendTextIfExist(result, self.processItem['parameter'], 'tolerance_up')
        result = _appendTextIfExist(result, self.processItem['parameter'], 'tolerance_down')
        result = _appendTextIfExist(result, self.processItem['parameter'], 'unit', True)
        return [] if result == '' else [result]
    
    @property
    def measurement(self):
        finalText = self.processItem['measurement']
        if self.processItem['readability'] != '' or self.processItem['parameter']['unit'] != '':
            finalText ='{} ({} {})'.format(
                finalText,
                self.processItem['readability'],
                self.processItem['parameter']['unit']
            )
        return self._splitLine(finalText)

    @property
    def controlMethodIntervalType(self):
        if self.processItem['control_method']['100_method'] == 'Auto check':
            return ['100%']
        return []

    @property
    def controlMethodIntervalPeriod(self):
        return self._splitLine(self.processItem['control_method']['interval'])
    
    @property
    def controlMethodSampleNo(self):
        if self.processItem['control_method']['sample_no'] > 1:
            return self._splitLine('(n={})'.format(self.processItem['control_method']['sample_no']))
        return []
    
    @property
    def controlMethodCalibrationInterval(self):
        return self._splitLine(self.processItem['control_method']['calibration_interval'])

    @property
    def controlMethodType(self):
        if self.processItem['control_method']['100_method'] == 'None':
            return []
        return self._splitLine(self.processItem['control_method']['100_method'])

    @property
    def controlMethodItemType(self):
        if self.processItem['control_method']['100_method'] == 'None':
            return self._splitLine(self.processItem['control_item_type'])
        return []

    @property
    def controlMethod(self):
        if self.processItem['control_method']['100_method'] == 'None':
            return ['Calibration']
        return []

    @property
    def controlMethodControlPerson(self):
        return self._splitLine(self.processItem['control_method']['in_charge'])

    @property
    def processCapability(self):
        textList = list()
        if self.processItem['initial_p_capability']['x_bar'] != '':
            textList.append('xbar = {}'.format(self.processItem['initial_p_capability']['x_bar']))
        if self.processItem['initial_p_capability']['cpk'] != '':
            textList.append('cpk = {}'.format(self.processItem['initial_p_capability']['cpk']))
        return textList

    @property
    def remark(self):
        return self._splitLine(self.processItem['remark']['remark'])

    @property
    def relateStandard(self):
        return self._splitLine(self.processItem['remark']['related_std'])

    @property
    def controlItemNo(self):
        return self.processItem['control_item_no']

    @property
    def controlItemType(self):
        return self.processItem['control_item_type']

    @property
    def checkTiming(self):
        return self.processItem['check_timing']

    def _splitLine(self, value: str):
        return value.split('\n')