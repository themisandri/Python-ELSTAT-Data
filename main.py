import streamlit as st
import pandas as pd
import urllib.request
import matplotlib.pyplot as plt
from data import get_tourists, per_country, per_month, per_vehicle

st.title("ELSTAT Data analysis")

st.write("Total tourists")
fig1 = get_tourists()
st.pyplot(fig1)

#clear current plot
plt.clf()

st.write("Total tourists by month")
fig_month = per_month()
st.pyplot(fig_month)

plt.clf()

st.write("Total tourists by country")
fig2 = per_country()
st.pyplot(fig2)

plt.clf()

st.write("Total tourists by vehicle")
fig4, fig5, fig6, fig7 = per_vehicle()
st.pyplot(fig4)
st.pyplot(fig5)
st.pyplot(fig6)
st.pyplot(fig7)





