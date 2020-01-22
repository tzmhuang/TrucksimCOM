h = actxserver('TruckSim.Application');
h.GoHome()

h.invoke()

ck = h.DataSetExists('','Baseline COM','External Control of Runs');

if ok
    return;
end

h.DeleteDataSet('','New Run Made with COM','External Control of Runs')

h.Gotolibrary('','Baseline COM','External Control of Runs')
h.CreateNew()
h.DatasetCategory('New Run Made with COM', 'External Control of Runs')

h.Checkbox('#Checkbox8','1')
h.Checkbox('#Checkbox3','1')
h.Ring('#RingCtrl10','1')
h.Yellow('*SPEED','150')

% h.Run('','')
% h.Checkbox('#Checkbox1','0')
% h.LaunchPlot()
