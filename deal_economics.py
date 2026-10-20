cols=st.columns(3 if kind=='Additive components' else 1)
for i,c in enumerate(cols):
    preferred=['H','HXS','SHS'][i] if kind=='Additive components' else 'H'
    idx=next((j for j,x in enumerate(options) if x[2]==preferred),0)
    selections.append(c.selectbox(f'Component {i+1}',options,index=idx,format_func=lambda x:f'{x[0]} | {x[1]} ({x[2]})'))
if len(set(selections))!=len(selections): st.error('Select distinct components.');st.stop()
st.subheader('Deal assumptions')
a,b,c,d=st.columns(4)
transport=a.number_input('Upstream transport ($/delivered MMBtu)',min_value=0.,value=defaults['transport'],step=.01,format='%.4f')
broker=b.number_input('Brokerage ($/delivered MMBtu)',min_value=0.,value=defaults['broker'],step=.01,format='%.4f')
margin=c.number_input('Target gross margin ($/MMBtu)',value=defaults['margin'],step=.01,format='%.4f')
fuel=d.number_input('Supply fuel retention (%)',min_value=0.,max_value=99.,value=0.,step=.1)/100
st.caption('Margin is after supply, transport and brokerage, after the capital charge, before overhead, other credit fees and taxes. Transport and brokerage apply to delivered usage. Fuel gross-up applies to gas supply only.')
st.subheader('Cost of capital')
cap1,cap2=st.columns(2)
capital_hurdle=cap1.number_input('Annual capital hurdle (%)',min_value=0.,value=8.,step=.5)/100
financing_days=cap2.number_input('Days of costs financed',min_value=0.,value=30.,step=1.)
st.caption('Capital charge per delivered MMBtu = positive supply + transport + brokerage cost per MMBtu × annual hurdle × financing days ÷ 365. The charge is included in the target offer and break-even price, and deducted from modeled profit. Financing days represent the cash collection gap. Negative total costs have a zero capital base.')
with st.expander('NPV and IRR assumptions'):
    a,b,c=st.columns(3)
    discount=a.number_input('Annual discount rate (%)',min_value=0.,value=10.,step=.5)/100
    investment=b.number_input('Initial cash investment ($)',min_value=0.,value=0.,step=1000.)
    recovery=c.number_input('Cash recovered at contract end ($)',min_value=0.,value=0.,step=1000.)
    st.caption('NPV is measured at contract start. Cash flows use month-end profit after the modeled capital charge, before overhead and taxes, less initial investment, plus terminal recovery. Enter posted collateral as initial cash and its expected return as terminal recovery. IRR is annualized and requires an initial cash outflow.')
months=pd.date_range(start,periods=term,freq='MS')
base=pd.DataFrame({'month':months,'usage_mmbtu':[defaults['usage'].get(m.month,0.) for m in months]})
st.subheader('Customer monthly usage')
st.caption('Enter total delivered MMBtu for each month. The attached workbook starts at 1,700 MMBtu/month. Values are estimates, not daily quantities.')
usage_file=st.file_uploader('Load a saved usage schedule (optional)',type=['csv'])
if usage_file is not None:
    try:
        loaded=pd.read_csv(usage_file)
        loaded['month']=pd.to_datetime(loaded['month']).dt.to_period('M').dt.to_timestamp()
        if loaded.month.duplicated().any(): raise ValueError('Duplicate usage months')
        base['usage_mmbtu']=base.month.map(loaded.set_index('month')['usage_mmbtu'])
        if base.usage_mmbtu.isna().any(): raise ValueError('Schedule must cover every selected month')
    except Exception as e: st.error(f'Invalid usage schedule: {e}');st.stop()
usage=st.data_editor(base,hide_index=True,disabled=['month'],column_config={'usage_mmbtu':st.column_config.NumberColumn('Usage (MMBtu/month)',min_value=0.,required=True)},key=f'usage_{start}_{term}',use_container_width=True)
manual=st.checkbox('Evaluate a proposed customer fixed price')
quote=st.number_input('Proposed customer price ($/MMBtu)',value=4.,step=.01) if manual else None
try:
    p=curve(prices,asof,selections,months)
    if p.isna().any().any():
        st.error('Pricing is missing for these delivery months/components. Select another trade date, product or term.')
        st.dataframe(p[p.isna().any(axis=1)]);st.stop()
    result,target=calculate(p,usage.usage_mmbtu,transport,broker,margin,fuel,quote,capital_hurdle,financing_days)
except ValueError as e: st.error(str(e));st.stop()
summary=deal_summary(result,discount,investment,recovery)
with summary_panel:
    st.subheader('Deal summary')
    st.caption(f'{customer} • Start {start_entered:%b %d, %Y} • {term} month term • Settlements {pd.Timestamp(asof):%b %d, %Y}')
    top=st.columns(4)
    top[0].metric('Annual revenue (annualized)',f"${summary['annual_revenue']:,.0f}")
    top[1].metric('Volume sold (contract MMBtu)',f"{summary['volume']:,.0f}")
    top[2].metric('Contract price ($/MMBtu)',f"${summary['contract_price']:,.4f}")
    top[3].metric('NPV at contract start',f"${summary['npv']:,.0f}")
    top[3].metric('Break-even price ($/MMBtu)',f'${target-margin:,.4f}')
    bottom=st.columns(3)
    bottom[0].metric('IRR (annualized)',f"{summary['irr']:.2%}" if summary['irr'] is not None else 'N/A')
    bottom[1].metric('MMBtu spread ($/MMBtu)',f"${summary['spread']:,.4f}")
    bottom[2].metric('Gross margin (%)',f"{summary['gross_margin']:.2%}" if summary['gross_margin'] is not None else 'N/A')
    st.caption(f"Capital hurdle: {capital_hurdle:.1%} • Financing days: {financing_days:g} • Weighted capital charge: ${(result.capital_cost.sum()/result.usage_mmbtu.sum()):.4f}/MMBtu • Contract capital charge: ${result.capital_cost.sum():,.0f}")
    st.caption('Annual revenue = total contract revenue × 12 ÷ term months. Spread = gross profit ÷ delivered volume. Gross margin = gross profit ÷ revenue. '+summary['irr_status']+'.')
    st.download_button('Download deal summary CSV',pd.DataFrame([dict(customer=customer,start_date=str(start_entered),term_months=term,trade_date=str(pd.Timestamp(asof).date()),capital_hurdle=capital_hurdle,financing_days=financing_days,capital_charge_per_mmbtu=result.capital_cost.sum()/result.usage_mmbtu.sum(),annual_discount_rate=discount,initial_investment=investment,terminal_recovery=recovery,**summary)]).to_csv(index=False).encode(),'deal_summary.csv','text/csv')
vol=result.usage_mmbtu.sum();gp=result.gross_profit.sum()
a,b,c,d=st.columns(4)
a.metric('Target customer fixed price',f'${target:,.4f}/MMBtu')
b.metric('Total contract volume',f'{vol:,.0f} MMBtu')
c.metric('Contract gross profit',f'${gp:,.0f}')
d.metric('Realized gross margin',f'${gp/vol:,.4f}/MMBtu')
st.caption(f'Customer: {customer} • Settlements: {pd.Timestamp(asof):%Y-%m-%d} • Annualized gross profit: ${gp*12/term:,.0f}')
st.subheader('Offer comparison')
st.caption('Offer uses the proposed customer price, or target price when no proposed price is entered. Costs and volume remain constant. Break-even includes the capital charge and excludes overhead and taxes.')
scenario_rows=[]
scenario_monthly=pd.DataFrame(index=result.index)
for label,reduction in [('Offer',0.),('Offer less $0.10',.10),('Offer less $0.20',.20)]:
    offer_price=float(result.customer_rate.iloc[0])-reduction
    sr,_=calculate(p,usage.usage_mmbtu,transport,broker,margin,fuel,offer_price,capital_hurdle,financing_days)
    ss=deal_summary(sr,discount,investment,recovery)
    scenario_rows.append({'Scenario':label,'Offer $/MMBtu':offer_price,'Break-even $/MMBtu':target-margin,'Spread $/MMBtu':ss['spread'],'Gross margin %':ss['gross_margin']*100 if ss['gross_margin'] is not None else None,'Average monthly GP $':sr.gross_profit.sum()/term,'Annualized GP $':sr.gross_profit.sum()*12/term,'Contract GP $':sr.gross_profit.sum(),'NPV $':ss['npv'],'Annualized IRR %':ss['irr']*100 if ss['irr'] is not None else None})
    scenario_monthly[label]=sr.gross_profit
scenario_table=pd.DataFrame(scenario_rows)
st.dataframe(scenario_table,hide_index=True,use_container_width=True)
st.caption('Monthly and annualized averages use the selected contract term. The schedule below shows actual monthly profits. NPV includes initial investment and recovery; break-even price is the volume-weighted cost including the capital charge.')
with st.expander('Monthly offer economics',expanded=True):
    st.line_chart(scenario_monthly)
    st.dataframe(scenario_monthly.reset_index(),hide_index=True,use_container_width=True)
st.download_button('Download offer comparison CSV',scenario_table.to_csv(index=False).encode(),'offer_comparison.csv','text/csv')
st.download_button('Download monthly offer economics CSV',scenario_monthly.reset_index().to_csv(index=False).encode(),'monthly_offer_economics.csv','text/csv')
t1,t2,t3=st.tabs(['Monthly economics','Term comparisons','Method & source'])
with t1:
    st.line_chart(result[['market_rate','landed_rate','all_in_cost_rate','customer_rate']])
    st.bar_chart(result[['gross_profit']])
    st.dataframe(result.reset_index(),hide_index=True,use_container_width=True)
    totals=result[['usage_mmbtu','revenue','supply_cost','transport_cost','broker_cost','capital_cost','gross_profit']].sum()
    st.write('Contract totals',totals.to_frame('Total').T)
    st.write('Selected-strip averages',pd.DataFrame({'Simple average':result[['market_rate','landed_rate','all_in_cost_rate','customer_rate']].mean(),'Volume-weighted average':result[['market_rate','landed_rate','all_in_cost_rate','customer_rate']].mul(result.usage_mmbtu,axis=0).sum()/vol}))
with t2:
    st.caption('Compares standard terms and your selected term within the entered usage schedule. Longer terms require extending the schedule above. Each term is priced separately to the same target margin.')
    rows=[]
    for n in sorted(set([1,3,6,12,18,24,36,48,term])):
        if n>term or usage.usage_mmbtu.iloc[:n].sum()<=0: continue
        r,q=calculate(p.iloc[:n],usage.usage_mmbtu.iloc[:n],transport,broker,margin,fuel,capital_hurdle=capital_hurdle,financing_days=financing_days)
        rows.append({'Months':n,'Volume MMBtu':r.usage_mmbtu.sum(),'Target fixed price':q,'Gross profit':r.gross_profit.sum(),'Annualized GP':r.gross_profit.sum()*12/n})
    st.dataframe(pd.DataFrame(rows),hide_index=True)
with t3:
    st.write('Supply = sum of selected components ÷ (1 − fuel retention). Landed cost = supply + upstream transport + brokerage. Target fixed price = volume-weighted landed cost + target margin. Capital charge = max(landed cost, 0) × annual hurdle × financing days / 365. Target fixed price = weighted landed cost + weighted capital charge + target margin. Monthly modeled profit = (customer fixed price − landed cost − capital charge) × usage.')
    st.write('Uses actual entered monthly volumes, calendar delivery months and a single settlement snapshot. No interpolation, silent zero filling or fallback to older settlements. Daily swing observations are not converted to forward prices.')
    st.write('The capital charge is a modeled financing expense, not a full working-capital schedule. NPV separately discounts profit after that charge; initial investment and terminal recovery are additional cash flows and should not duplicate the financed operating costs. The workbook’s NPV/carry formulas and daily-volume approximation are not reproduced. This model quotes a level fixed price over the selected term. It excludes utility charges, imbalance exposure, credit/collateral financing and overhead unless included in your rates.')
    st.dataframe(pd.DataFrame(selections,columns=['Hub','Product','Contract']),hide_index=True)
export=result.reset_index(); export.insert(0,'customer',customer);export.insert(1,'trade_date',pd.Timestamp(asof));export.insert(2,'contract_start_date',str(start_entered));export.insert(3,'term_months',term);export.insert(2,'components',' + '.join(x[2] for x in selections))
st.download_button('Download monthly economics CSV',export.to_csv(index=False).encode(),'rem_deal_economics.csv','text/csv')
st.download_button('Download usage schedule CSV',usage.to_csv(index=False).encode(),'customer_usage.csv','text/csv')
